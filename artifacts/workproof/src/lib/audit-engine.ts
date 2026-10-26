import * as pdfjs from "pdfjs-dist";
import pdfWorkerUrl from "pdfjs-dist/build/pdf.worker.min.mjs?url";
import mammoth from "mammoth";
import Papa from "papaparse";
import * as XLSX from "xlsx";

pdfjs.GlobalWorkerOptions.workerSrc = pdfWorkerUrl;

export type UploadKind = "contract" | "communication" | "work";
export type Severity = "Critical" | "High" | "Medium" | "Low";
export type FindingCategory =
  | "Unbilled work"
  | "Scope creep"
  | "Missing approval"
  | "Missing evidence"
  | "Over-delivery"
  | "Missed commitment"
  | "Billing mismatch";
export type CommitmentStatus =
  | "Open"
  | "In progress"
  | "Completed"
  | "Overdue"
  | "At risk"
  | "Unverified";

export interface Evidence {
  fileName: string;
  excerpt: string;
  lineNumber?: number;
  kind: UploadKind;
}

export interface ParsedUpload {
  id: string;
  name: string;
  kind: UploadKind;
  text: string;
  rows: string[][];
  uploadedAt: string;
}

export interface ScopeRow {
  work: string;
  contract: number | null;
  actual: number | null;
  difference: number | null;
  unitPrice: number | null;
  status: "within-scope" | "over-scope" | "unknown";
  contractEvidence?: Evidence;
  actualEvidence?: Evidence;
}

export interface TimelineEvent {
  date: string;
  label: string;
  description: string;
  evidence?: Evidence;
}

export interface Finding {
  id: string;
  severity: Severity;
  category: FindingCategory;
  title: string;
  client: string;
  estimatedValue: number | null;
  calculation: string | null;
  explanation: string;
  sources: string[];
  recommendedAction: string;
  confidence: number;
  approvalStatus: string;
  evidence: Evidence[];
  chronology: TimelineEvent[];
  originalScope: number | null;
  additionalWork: number | null;
  revenueAtRisk: number | null;
  status: "Open" | "Resolved" | "Dismissed";
}

export interface Commitment {
  id: string;
  client: string;
  commitment: string;
  owner: string;
  dueDate: string | null;
  source: string;
  status: CommitmentStatus;
  evidence: Evidence | null;
  estimatedValue: number | null;
}

export interface AuditReport {
  id: string;
  project: string;
  client: string;
  createdAt: string;
  isDemo: boolean;
  files: ParsedUpload[];
  findings: Finding[];
  commitments: Commitment[];
  scopeRows: ScopeRow[];
  revenueAtRisk: number | null;
  potentialUnbilled: number | null;
  insufficientData: boolean;
  determinations: string[];
  cannotDetermine: string[];
}

export class FileProcessingError extends Error {
  constructor(message = "Could not process this file.") {
    super(message);
    this.name = "FileProcessingError";
  }
}

const allowedExtensions: Record<UploadKind, string[]> = {
  contract: [".pdf", ".docx", ".txt"],
  communication: [".txt", ".docx", ".pdf", ".csv"],
  work: [".csv", ".xlsx"],
};

const workColumnPattern =
  /\b(task|work|deliverable|description|item|activity|title|service|project task)\b/i;
const communicationColumnPattern =
  /\b(message|body|content|communication|subject|email|transcript|notes|text|chat|comment)\b/i;
const quantityColumnPattern = /\b(quantity|qty|count|units|number completed)\b/i;

function hasNonEmptyDataRows(rows: string[][]) {
  return rows.slice(1).some((row) => row.some((cell) => String(cell).trim()));
}

function validateTabularRows(rows: string[][], kind: UploadKind): string[] {
  const header = rows[0]?.map((cell) => String(cell).trim()) ?? [];
  const validHeader =
    kind === "work"
      ? header.some((cell) => workColumnPattern.test(cell))
      : header.some((cell) => communicationColumnPattern.test(cell));
  if (!validHeader || !hasNonEmptyDataRows(rows)) {
    const expected =
      kind === "work"
        ? "a task, work, deliverable, description, or activity column and at least one data row"
        : "a message, body, content, subject, email, or transcript column and at least one data row";
    throw new FileProcessingError(
      `Could not process this file. The ${kind} spreadsheet needs ${expected}.`,
    );
  }
  return header;
}

export function getAllowedExtensions(kind: UploadKind) {
  return allowedExtensions[kind].join(",");
}

export async function parseUploadedFile(
  file: File,
  kind: UploadKind,
): Promise<ParsedUpload> {
  const extension = `.${file.name.split(".").pop()?.toLowerCase() ?? ""}`;
  if (!allowedExtensions[kind].includes(extension)) {
    throw new FileProcessingError(
      `Could not process this file. ${extension || "This file type"} is not supported for ${kind}.`,
    );
  }
  if (file.size === 0) {
    throw new FileProcessingError();
  }

  try {
    let text = "";
    let rows: string[][] = [];

    if (extension === ".txt") {
      text = await file.text();
      rows = text.split(/\r?\n/).map((line) => [line]);
    } else if (extension === ".csv") {
      const csv = Papa.parse<string[]>(await file.text(), {
        skipEmptyLines: true,
      });
      if (csv.errors.length) {
        throw new Error(csv.errors[0]?.message);
      }
      rows = csv.data;
      if (kind === "work" || kind === "communication") {
        validateTabularRows(rows, kind);
      }
      text = rows.map((row) => row.join(" | ")).join("\n");
    } else if (extension === ".xlsx") {
      const workbook = XLSX.read(await file.arrayBuffer(), { type: "array" });
      const parsedSheets = workbook.SheetNames.map((sheetName) => {
        const sheet = workbook.Sheets[sheetName];
        const sheetRows = XLSX.utils.sheet_to_json<string[]>(sheet, {
          header: 1,
          raw: false,
          defval: "",
        });
        return { sheetName, rows: sheetRows };
      });
      const validSheets = parsedSheets.filter(({ rows: sheetRows }) => {
        try {
          validateTabularRows(sheetRows, kind);
          return true;
        } catch {
          return false;
        }
      });
      if (validSheets.length === 0) validateTabularRows(parsedSheets[0]?.rows ?? [], kind);
      rows = validSheets.flatMap(({ rows: sheetRows }) => sheetRows);
      text = validSheets
        .map(
          ({ sheetName, rows: sheetRows }) =>
            `[${sheetName}]\n${sheetRows.map((row) => row.join(" | ")).join("\n")}`,
        )
        .join("\n\n");
    } else if (extension === ".docx") {
      const result = await mammoth.extractRawText({
        arrayBuffer: await file.arrayBuffer(),
      });
      text = result.value;
      rows = text.split(/\r?\n/).map((line) => [line]);
    } else if (extension === ".pdf") {
      const pdf = await pdfjs.getDocument({
        data: await file.arrayBuffer(),
      }).promise;
      const pages: string[] = [];
      for (let pageNumber = 1; pageNumber <= pdf.numPages; pageNumber += 1) {
        const page = await pdf.getPage(pageNumber);
        const content = await page.getTextContent();
        const pageText = content.items
          .map((item) => ("str" in item ? item.str : ""))
          .join(" ")
          .trim();
        if (pageText) pages.push(`[Page ${pageNumber}]\n${pageText}`);
      }
      text = pages.join("\n\n");
      rows = text.split(/\r?\n/).map((line) => [line]);
    }

    if (!text.trim()) throw new Error("No readable text was found.");
    return {
      id: crypto.randomUUID(),
      name: file.name,
      kind,
      text,
      rows,
      uploadedAt: new Date().toISOString(),
    };
  } catch (error) {
    if (error instanceof FileProcessingError) throw error;
    throw new FileProcessingError();
  }
}

interface DeliverableRule {
  key: string;
  label: string;
  aliases: RegExp;
}

const deliverableRules: DeliverableRule[] = [
  {
    key: "landing-pages",
    label: "Landing pages",
    aliases: /\blanding[\s-]?pages?\b|\bLPs?\b/i,
  },
  {
    key: "blog-posts",
    label: "Blog posts",
    aliases: /\bblog\s+posts?\b|\barticles?\b/i,
  },
  {
    key: "social-creatives",
    label: "Social creatives",
    aliases: /\bsocial\s+(?:media\s+)?creatives?\b|\bsocial\s+posts?\b/i,
  },
];

function evidenceForLine(file: ParsedUpload, line: string, index: number): Evidence {
  return {
    fileName: file.name,
    excerpt: line.trim().slice(0, 1500),
    lineNumber: index + 1,
    kind: file.kind,
  };
}

function linesOf(files: ParsedUpload | ParsedUpload[]) {
  const list = Array.isArray(files) ? files : [files];
  return list.flatMap((file) =>
    file.text
      .split(/\r?\n/)
      .map((line, index) => ({ file, line: line.trim(), index }))
      .filter(({ line }) => line.length > 0),
  );
}

function findQuantity(line: string, rule: DeliverableRule): number | null {
  const match = line.match(
    new RegExp(`(?:^|\\b)(\\d+(?:\\.\\d+)?)\\s+${rule.aliases.source}`, "i"),
  );
  if (match?.[1]) return Number(match[1]);

  const reverseMatch = line.match(
    new RegExp(
      `${rule.aliases.source}\\s*(?:[:=|,-]|\\bcount\\b)\\s*(\\d+(?:\\.\\d+)?)`,
      "i",
    ),
  );
  return reverseMatch?.[1] ? Number(reverseMatch[1]) : null;
}

function findExplicitUnitPrice(line: string, rule: DeliverableRule) {
  const patterns = [
    new RegExp(
      `\\$\\s*(\\d+(?:,\\d{3})*(?:\\.\\d{1,2})?)\\s*(?:\\/\\s*(?:${rule.aliases.source}|(?:page|post|creative|unit))|per\\s+(?:${rule.aliases.source}|(?:page|post|creative|unit))|each\\s+(?:${rule.aliases.source}))`,
      "i",
    ),
    new RegExp(
      `(?:${rule.aliases.source})[^\\n$]{0,60}\\$\\s*(\\d+(?:,\\d{3})*(?:\\.\\d{1,2})?)\\s*(?:each|per\\s+(?:page|post|creative|unit)|\\/\\s*(?:page|post|creative|unit))`,
      "i",
    ),
    new RegExp(
      `(?:${rule.aliases.source})[^\\n$]{0,60}\\b(?:unit\\s+rate|rate\\s+per\\s+unit)\\s*[:=]?\\s*\\$\\s*(\\d+(?:,\\d{3})*(?:\\.\\d{1,2})?)`,
      "i",
    ),
  ];
  for (const pattern of patterns) {
    const match = line.match(pattern);
    if (match?.[1]) return Number(match[1].replaceAll(",", ""));
  }
  return null;
}

function buildScopeRows(
  contractFiles: ParsedUpload[],
  workFiles: ParsedUpload[],
): ScopeRow[] {
  if (contractFiles.length === 0 || workFiles.length === 0) return [];
  const contractLines = linesOf(contractFiles);
  const workLines = linesOf(workFiles);

  return deliverableRules.flatMap((rule) => {
    const contractMatch = contractLines.find(({ line }) =>
      rule.aliases.test(line),
    );
    const contractQuantity = contractMatch
      ? findQuantity(contractMatch.line, rule)
      : null;
    const actualMatches = workLines.filter(({ line }) =>
      rule.aliases.test(line),
    );
    if (contractQuantity === null && actualMatches.length === 0) return [];

    const explicitWorkQuantities = actualMatches
      .map(({ line }) => findQuantity(line, rule))
      .filter((value): value is number => value !== null);
    const actualQuantity =
      actualMatches.length > 0
        ? explicitWorkQuantities.length > 0
          ? explicitWorkQuantities.reduce((sum, value) => sum + value, 0)
          : actualMatches.length
        : null;
    const explicitRate = contractLines.find(({ line }) =>
      findExplicitUnitPrice(line, rule),
    );
    const unitPrice = explicitRate
      ? findExplicitUnitPrice(explicitRate.line, rule)
      : null;
    const difference =
      contractQuantity !== null && actualQuantity !== null
        ? actualQuantity - contractQuantity
        : null;

    return [
      {
        work: rule.label,
        contract: contractQuantity,
        actual: actualQuantity,
        difference,
        unitPrice,
        status:
          difference === null
            ? "unknown"
            : difference > 0
              ? "over-scope"
              : "within-scope",
        contractEvidence: contractMatch
          ? evidenceForLine(
              contractMatch.file,
              contractMatch.line,
              contractMatch.index,
            )
          : undefined,
        actualEvidence: actualMatches[0]
          ? evidenceForLine(
              actualMatches[0].file,
              actualMatches[0].line,
              actualMatches[0].index,
            )
          : undefined,
      },
    ];
  });
}

function findRequestEvidence(
  communicationFiles: ParsedUpload[],
  work: ScopeRow,
) {
  if (communicationFiles.length === 0) return [];
  const requestPattern =
    /\b(?:request(?:ed)?|please|can you|could you|add|additional|extra|another|also|need|would like|create|build|design)\b/i;
  return linesOf(communicationFiles)
    .filter(({ line }) => workLabelMatches(line, work.work) && requestPattern.test(line))
    .slice(0, 2)
    .map(({ file, line, index }) =>
      evidenceForLine(file, line, index),
    );
}

function workLabelMatches(line: string, label: string) {
  const rule = deliverableRules.find((candidate) => candidate.label === label);
  return rule?.aliases.test(line) ?? false;
}

function amountFromContractValue(scopeRow: ScopeRow, extraQuantity: number) {
  if (scopeRow.unitPrice !== null) return scopeRow.unitPrice * extraQuantity;
  return null;
}

export function calculateRevenueAtRisk(
  findings: Pick<Finding, "estimatedValue">[],
): number | null {
  if (
    findings.some(
      ({ estimatedValue }) =>
        estimatedValue === null ||
        !Number.isFinite(estimatedValue) ||
        estimatedValue < 0,
    )
  ) {
    return null;
  }
  return findings.reduce(
    (total, finding) => total + (finding.estimatedValue ?? 0),
    0,
  );
}

function inferClient(contractFiles: ParsedUpload[], fallback: string) {
  const patterns = [
    /(?:client|prepared for|customer)\s*[:\-]\s*([A-Z][A-Za-z0-9& .'-]{2,50})/i,
    /(?:agreement|statement of work)\s+(?:with|for)\s+([A-Z][A-Za-z0-9& .'-]{2,50})/i,
  ];
  for (const file of contractFiles) {
    for (const pattern of patterns) {
      const match = file.text.match(pattern);
      if (match?.[1]) return match[1].trim().replace(/[.,;]+$/, "");
    }
  }
  return fallback || "Unknown client";
}

function extractCommitments(
  communicationFiles: ParsedUpload[],
  client: string,
): Commitment[] {
  if (communicationFiles.length === 0) return [];
  const commitmentPattern =
    /\b(?:we(?:'ll| will)|i(?:'ll| will)|committed to|promise(?:d)? to|deliver|complete|send|share|provide|finish)\b/i;
  return linesOf(communicationFiles)
    .filter(({ line }) => commitmentPattern.test(line))
    .slice(0, 20)
    .map(({ file, line, index }, position) => {
      const dueMatch = line.match(
        /\b(?:by|before|due(?: date)?[: ]*)\s+((?:jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\s+\d{1,2}(?:,?\s+\d{4})?|\d{1,2}[/-]\d{1,2}(?:[/-]\d{2,4})?)/i,
      );
      const evidence = evidenceForLine(file, line, index);
      return {
        id: `commitment-${position + 1}`,
        client,
        commitment: line,
        owner: "Not available",
        dueDate: dueMatch?.[1] ?? null,
        source: file.name,
        status: dueMatch ? "Open" : "Unverified",
        evidence,
        estimatedValue: null,
      };
    });
}

export function analyzeProject(
  files: ParsedUpload[],
  projectName: string,
): AuditReport {
  const uniqueFiles = files.filter((file, index) => {
    const signature = `${file.kind}\u0000${file.text.replace(/\s+/g, " ").trim()}`;
    return files.findIndex((candidate) =>
      `${candidate.kind}\u0000${candidate.text.replace(/\s+/g, " ").trim()}` === signature,
    ) === index;
  });
  const contractFiles = uniqueFiles.filter((file) => file.kind === "contract");
  const communicationFiles = uniqueFiles.filter(
    (file) => file.kind === "communication",
  );
  const workFiles = uniqueFiles.filter((file) => file.kind === "work");
  const client = inferClient(contractFiles, projectName);
  const scopeRows = buildScopeRows(contractFiles, workFiles);
  const findings: Finding[] = [];

  scopeRows.forEach((row, index) => {
    if (row.difference === null || row.difference <= 0) return;
    const requestEvidence = findRequestEvidence(communicationFiles, row);
    const evidence = [
      ...(row.contractEvidence ? [row.contractEvidence] : []),
      ...(row.actualEvidence ? [row.actualEvidence] : []),
      ...requestEvidence,
    ];
    const estimatedValue = amountFromContractValue(row, row.difference);
    const hasRequest = requestEvidence.length > 0;
    const title = `${row.difference} additional ${row.work.toLowerCase()} detected outside SOW`;
    const calculation =
      estimatedValue !== null && row.unitPrice !== null
        ? `${row.difference} × $${row.unitPrice.toLocaleString("en-US")} per unit = $${estimatedValue.toLocaleString("en-US")}`
        : null;
    findings.push({
      id: `finding-scope-${index + 1}`,
      severity: hasRequest ? "High" : "Medium",
      category: "Scope creep",
      title,
      client,
      estimatedValue,
      calculation,
      explanation: `${row.contract ?? "Unspecified"} ${row.work.toLowerCase()} in the contract and ${row.actual} reflected in work data. ${hasRequest ? "A related client request appears in the communication file." : "No matching request evidence was found in the communication file."}`,
      sources: [...new Set(evidence.map((item) => item.fileName))],
      recommendedAction: hasRequest
        ? "Confirm scope approval and create a change order."
        : "Review the work record against the agreed scope before billing.",
      confidence: Math.round(
        ((row.contractEvidence ? 1 : 0) +
          (row.actualEvidence ? 1 : 0) +
          (requestEvidence.length > 0 ? 1 : 0)) /
          3 *
          100,
      ),
      approvalStatus: "Formal approval not verified in provided files",
      evidence,
      chronology: [
        ...(requestEvidence[0]
          ? [
              {
                date: "From communication file",
                label: "Request found",
                description: requestEvidence[0].excerpt,
                evidence: requestEvidence[0],
              },
            ]
          : []),
        ...(row.actualEvidence
          ? [
              {
                date: "From work file",
                label: "Work recorded",
                description: row.actualEvidence.excerpt,
                evidence: row.actualEvidence,
              },
            ]
          : []),
        {
          date: "Audit result",
          label: "Approval status",
          description:
            "Formal approval cannot be confirmed from the uploaded files.",
        },
      ],
      originalScope: row.contract,
      additionalWork: row.difference,
      revenueAtRisk: estimatedValue,
      status: "Open",
    });
  });

  const commitments = extractCommitments(communicationFiles, client);
  const hasMissingFiles =
    contractFiles.length === 0 ||
    workFiles.length === 0 ||
    communicationFiles.length === 0;
  const hasUnpricedRisk = findings.some(
    (finding) => finding.revenueAtRisk === null,
  );
  const determinations: string[] = [];
  const cannotDetermine: string[] = [];
  if (contractFiles.length > 0) {
    determinations.push("Contract text was extracted from the uploaded file(s).");
  } else {
    cannotDetermine.push("Contract scope and agreed quantities: contract file not provided.");
  }
  if (communicationFiles.length > 0) {
    determinations.push("Communication text was extracted and checked for requests and commitments.");
  } else {
    cannotDetermine.push("Client requests and commitments: communication file not provided.");
  }
  if (workFiles.length > 0) {
    determinations.push("Work data rows were extracted from the uploaded file(s).");
  } else {
    cannotDetermine.push("Completed work quantities: work data file not provided.");
  }
  if (scopeRows.length > 0) {
    determinations.push(`Contract quantities were compared with recorded work for ${scopeRows.length} deliverable type(s).`);
  } else {
    cannotDetermine.push("Scope comparison: matching contract and work records were not both available.");
  }
  cannotDetermine.push("Invoice matching: no invoice records were included in the upload flow.");
  if (findings.some((finding) => finding.estimatedValue === null)) {
    cannotDetermine.push("Revenue value for one or more deviations: a verified unit rate was not found.");
  }
  if (findings.length > 0) {
    cannotDetermine.push("Formal approval: approval evidence was not verified in the provided files.");
  }
  return {
    id: crypto.randomUUID(),
    project: projectName || client,
    client,
    createdAt: new Date().toISOString(),
    isDemo: false,
    files: uniqueFiles,
    findings,
    commitments,
    scopeRows,
    revenueAtRisk:
      hasMissingFiles || hasUnpricedRisk
        ? null
        : calculateRevenueAtRisk(findings),
    potentialUnbilled: null,
    insufficientData: hasUnpricedRisk || hasMissingFiles || findings.length === 0,
    determinations,
    cannotDetermine,
  };
}

export const demoReport: AuditReport = {
  id: "demo-northstar-media",
  project: "Monthly growth retainer",
  client: "Northstar Media",
  createdAt: "2026-10-25T09:00:00.000Z",
  isDemo: true,
  files: [],
  revenueAtRisk: 2450,
  potentialUnbilled: 2450,
  insufficientData: false,
  determinations: [
    "DEMO DATA · The Northstar example compares the stated agreement scope with illustrative communication and work records.",
  ],
  cannotDetermine: [
    "DEMO DATA · Invoice matching is not assessed because invoice records are not part of the sample.",
  ],
  scopeRows: [
    {
      work: "Landing pages",
      contract: 4,
      actual: 6,
      difference: 2,
      unitPrice: null,
      status: "over-scope",
    },
    {
      work: "Blog posts",
      contract: 12,
      actual: 12,
      difference: 0,
      unitPrice: null,
      status: "within-scope",
    },
    {
      work: "Social creatives",
      contract: 20,
      actual: 28,
      difference: 8,
      unitPrice: null,
      status: "over-scope",
    },
  ],
  findings: [
    {
      id: "demo-finding-landing-pages",
      severity: "High",
      category: "Scope creep",
      title: "2 additional landing pages completed",
      client: "Northstar Media",
      estimatedValue: 1600,
      calculation: "DEMO DATA · illustrative allocation toward the preset $2,450 demo total; no unit rate is claimed.",
      explanation:
        "The monthly SOW includes 4 landing pages. A client request for two more appears in the communication sample, and work data shows 6 pages completed. Demo estimate based on the sample's stated change-order value.",
      sources: ["Northstar_SOW_demo.txt", "Client_thread_demo.txt", "Work_log_demo.csv"],
      recommendedAction: "Create Change Order",
      confidence: 94,
      approvalStatus: "No formal scope approval detected in this demo sample",
      evidence: [
        {
          fileName: "Northstar_SOW_demo.txt",
          excerpt: "Landing pages: 4 per month",
          lineNumber: 7,
          kind: "contract",
        },
        {
          fileName: "Client_thread_demo.txt",
          excerpt:
            "Could we add two landing pages for the campaign next week?",
          lineNumber: 12,
          kind: "communication",
        },
        {
          fileName: "Work_log_demo.csv",
          excerpt: "Landing pages | 6 | completed",
          lineNumber: 8,
          kind: "work",
        },
      ],
      chronology: [
        {
          date: "Oct 18",
          label: "Client request",
          description: "Requested two additional landing pages.",
        },
        {
          date: "Oct 19",
          label: "Agency acknowledgment",
          description: "Agency confirmed it could deliver the campaign pages.",
        },
        {
          date: "Oct 21",
          label: "Work recorded",
          description: "The work log shows six pages completed against four scoped.",
        },
        {
          date: "Oct 25",
          label: "Approval and invoice check",
          description:
            "No formal approval or invoice record was included in this demo sample.",
        },
      ],
      originalScope: 4,
      additionalWork: 2,
      revenueAtRisk: 1600,
      status: "Open",
    },
    {
      id: "demo-finding-social",
      severity: "Medium",
      category: "Over-delivery",
      title: "8 extra social creatives recorded",
      client: "Northstar Media",
      estimatedValue: 850,
      calculation: "DEMO DATA · illustrative allocation toward the preset $2,450 demo total; no unit rate is claimed.",
      explanation:
        "The SOW includes 20 social creatives and the demo work log shows 28 completed. The demo estimate is allocated from the sample's stated $2,450 total risk.",
      sources: ["Northstar_SOW_demo.txt", "Work_log_demo.csv"],
      recommendedAction: "Review extra deliverables and confirm change-order status.",
      confidence: 89,
      approvalStatus: "Approval not verified in this demo sample",
      evidence: [
        {
          fileName: "Northstar_SOW_demo.txt",
          excerpt: "Social creatives: 20 per month",
          lineNumber: 9,
          kind: "contract",
        },
        {
          fileName: "Work_log_demo.csv",
          excerpt: "Social creatives | 28 | completed",
          lineNumber: 14,
          kind: "work",
        },
      ],
      chronology: [
        {
          date: "Oct 21",
          label: "Work recorded",
          description: "The work log shows 28 social creatives completed.",
        },
        {
          date: "Oct 25",
          label: "Scope check",
          description: "The contract sample includes 20 social creatives.",
        },
      ],
      originalScope: 20,
      additionalWork: 8,
      revenueAtRisk: 850,
      status: "Open",
    },
  ],
  commitments: [
    {
      id: "demo-commitment-1",
      client: "Northstar Media",
      commitment: "Deliver 4 landing pages for the campaign",
      owner: "Agency team",
      dueDate: "Oct 25, 2026",
      source: "Northstar_SOW_demo.txt",
      status: "Completed",
      evidence: {
        fileName: "Work_log_demo.csv",
        excerpt: "Landing pages | 6 | completed",
        lineNumber: 8,
        kind: "work",
      },
      estimatedValue: null,
    },
    {
      id: "demo-commitment-2",
      client: "Northstar Media",
      commitment: "Publish 12 blog posts during the retainer month",
      owner: "Agency team",
      dueDate: null,
      source: "Northstar_SOW_demo.txt",
      status: "Completed",
      evidence: {
        fileName: "Work_log_demo.csv",
        excerpt: "Blog posts | 12 | completed",
        lineNumber: 11,
        kind: "work",
      },
      estimatedValue: null,
    },
  ],
};

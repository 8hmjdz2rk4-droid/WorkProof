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
      if (csv.errors.length && csv.data.length === 0) {
        throw new Error(csv.errors[0]?.message);
      }
      rows = csv.data;
      text = rows.map((row) => row.join(" | ")).join("\n");
    } else if (extension === ".xlsx") {
      const workbook = XLSX.read(await file.arrayBuffer(), { type: "array" });
      rows = workbook.SheetNames.flatMap((sheetName) => {
        const sheet = workbook.Sheets[sheetName];
        const sheetRows = XLSX.utils.sheet_to_json<string[]>(sheet, {
          header: 1,
          raw: false,
          defval: "",
        });
        text += `\n[${sheetName}]\n${sheetRows.map((row) => row.join(" | ")).join("\n")}`;
        return sheetRows;
      });
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

function evidenceForLine(
  file: ParsedUpload,
  line: string,
  index: number,
): Evidence {
  return {
    fileName: file.name,
    excerpt: line.trim(),
    lineNumber: index + 1,
    kind: file.kind,
  };
}

function linesOf(file: ParsedUpload) {
  return file.text
    .split(/\r?\n/)
    .map((line, index) => ({ line: line.trim(), index }))
    .filter(({ line }) => line.length > 0);
}

function findQuantity(line: string, rule: DeliverableRule): number | null {
  const match = line.match(
    new RegExp(`(?:^|\\b)(\\d+(?:\\.\\d+)?)\\s+${rule.aliases.source}`, "i"),
  );
  if (match?.[1]) return Number(match[1]);

  const reverseMatch = line.match(
    new RegExp(`${rule.aliases.source}[^\\d$]{0,30}(\\d+(?:\\.\\d+)?)`, "i"),
  );
  return reverseMatch?.[1] ? Number(reverseMatch[1]) : null;
}

function findExplicitUnitPrice(line: string, rule: DeliverableRule) {
  const patterns = [
    new RegExp(
      `\\$\\s*(\\d+(?:,\\d{3})*(?:\\.\\d{1,2})?)\\s*(?:per|each|/)\\s*(?:${rule.aliases.source})`,
      "i",
    ),
    new RegExp(
      `(?:${rule.aliases.source})[^\\n$]{0,60}\\$\\s*(\\d+(?:,\\d{3})*(?:\\.\\d{1,2})?)\\s*(?:each|per\\s+(?:page|post|creative|unit))`,
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
  contractFile: ParsedUpload | undefined,
  workFile: ParsedUpload | undefined,
): ScopeRow[] {
  if (!contractFile || !workFile) return [];
  const contractLines = linesOf(contractFile);
  const workLines = linesOf(workFile);

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
    const unitPrice = contractMatch
      ? findExplicitUnitPrice(contractMatch.line, rule)
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
              contractFile,
              contractMatch.line,
              contractMatch.index,
            )
          : undefined,
        actualEvidence: actualMatches[0]
          ? evidenceForLine(workFile, actualMatches[0].line, actualMatches[0].index)
          : undefined,
      },
    ];
  });
}

function findRequestEvidence(
  communicationFile: ParsedUpload | undefined,
  work: ScopeRow,
) {
  if (!communicationFile) return [];
  const requestPattern =
    /\b(?:request(?:ed)?|please|can you|could you|add|additional|extra|another|also|need|would like|create|build|design)\b/i;
  return linesOf(communicationFile)
    .filter(({ line }) => workLabelMatches(line, work.work) && requestPattern.test(line))
    .slice(0, 2)
    .map(({ line, index }) =>
      evidenceForLine(communicationFile, line, index),
    );
}

function workLabelMatches(line: string, label: string) {
  const rule = deliverableRules.find((candidate) => candidate.label === label);
  return rule?.aliases.test(line) ?? false;
}

function amountFromContractValue(
  contractFile: ParsedUpload | undefined,
  scopeRow: ScopeRow,
  extraQuantity: number,
) {
  if (scopeRow.unitPrice !== null) return scopeRow.unitPrice * extraQuantity;
  if (!contractFile) return null;
  const explicitRateLines = linesOf(contractFile).filter(({ line }) =>
    findExplicitUnitPrice(
      line,
      deliverableRules.find((rule) => rule.label === scopeRow.work)!,
    ),
  );
  if (explicitRateLines.length > 0) {
    const rate = findExplicitUnitPrice(
      explicitRateLines[0].line,
      deliverableRules.find((rule) => rule.label === scopeRow.work)!,
    );
    if (rate !== null) return rate * extraQuantity;
  }
  return null;
}

function inferClient(contractFile: ParsedUpload | undefined, fallback: string) {
  if (!contractFile) return fallback || "Unknown client";
  const text = contractFile.text;
  const patterns = [
    /(?:client|prepared for|customer)\s*[:\-]\s*([A-Z][A-Za-z0-9& .'-]{2,50})/i,
    /(?:agreement|statement of work)\s+(?:with|for)\s+([A-Z][A-Za-z0-9& .'-]{2,50})/i,
  ];
  for (const pattern of patterns) {
    const match = text.match(pattern);
    if (match?.[1]) return match[1].trim().replace(/[.,;]+$/, "");
  }
  return fallback || "Unknown client";
}

function extractCommitments(
  communicationFile: ParsedUpload | undefined,
  client: string,
): Commitment[] {
  if (!communicationFile) return [];
  const commitmentPattern =
    /\b(?:we(?:'ll| will)|i(?:'ll| will)|committed to|promise(?:d)? to|deliver|complete|send|share|provide|finish)\b/i;
  return linesOf(communicationFile)
    .filter(({ line }) => commitmentPattern.test(line))
    .slice(0, 20)
    .map(({ line, index }, position) => {
      const dueMatch = line.match(
        /\b(?:by|before|due(?: date)?[: ]*)\s+((?:jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)[a-z]*\s+\d{1,2}(?:,?\s+\d{4})?|\d{1,2}[/-]\d{1,2}(?:[/-]\d{2,4})?)/i,
      );
      const evidence = evidenceForLine(communicationFile, line, index);
      return {
        id: `commitment-${position + 1}`,
        client,
        commitment: line,
        owner: "Unassigned",
        dueDate: dueMatch?.[1] ?? null,
        source: communicationFile.name,
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
  const contractFile = files.find((file) => file.kind === "contract");
  const communicationFile = files.find((file) => file.kind === "communication");
  const workFile = files.find((file) => file.kind === "work");
  const client = inferClient(contractFile, projectName);
  const scopeRows = buildScopeRows(contractFile, workFile);
  const findings: Finding[] = [];

  scopeRows.forEach((row, index) => {
    if (row.difference === null || row.difference <= 0) return;
    const requestEvidence = findRequestEvidence(communicationFile, row);
    const evidence = [
      ...(row.contractEvidence ? [row.contractEvidence] : []),
      ...(row.actualEvidence ? [row.actualEvidence] : []),
      ...requestEvidence,
    ];
    const estimatedValue = amountFromContractValue(
      contractFile,
      row,
      row.difference,
    );
    const hasRequest = requestEvidence.length > 0;
    const title = `${row.difference} additional ${row.work.toLowerCase()} detected outside SOW`;
    findings.push({
      id: `finding-scope-${index + 1}`,
      severity: hasRequest ? "High" : "Medium",
      category: "Scope creep",
      title,
      client,
      estimatedValue,
      explanation: `${row.contract ?? "Unspecified"} ${row.work.toLowerCase()} in the contract and ${row.actual} reflected in work data. ${hasRequest ? "A related client request appears in the communication file." : "No matching request evidence was found in the communication file."}`,
      sources: [...new Set(evidence.map((item) => item.fileName))],
      recommendedAction: hasRequest
        ? "Confirm scope approval and create a change order."
        : "Review the work record against the agreed scope before billing.",
      confidence: hasRequest ? 92 : 78,
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

  const commitments = extractCommitments(communicationFile, client);
  const hasMissingFiles = !contractFile || !workFile || !communicationFile;
  const pricedRisk = findings.reduce(
    (sum, finding) => sum + (finding.revenueAtRisk ?? 0),
    0,
  );
  const hasUnpricedRisk = findings.some(
    (finding) => finding.revenueAtRisk === null,
  );
  return {
    id: crypto.randomUUID(),
    project: projectName || client,
    client,
    createdAt: new Date().toISOString(),
    isDemo: false,
    files,
    findings,
    commitments,
    scopeRows,
    revenueAtRisk: hasUnpricedRisk || hasMissingFiles ? null : pricedRisk,
    potentialUnbilled: null,
    insufficientData: hasUnpricedRisk || hasMissingFiles || findings.length === 0,
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

import {
  demoReport,
  type AuditReport,
  type FindingCategory,
} from "./audit-engine";

const storageKey = "workproof.audit-reports.v1";

function readStoredReports(): AuditReport[] {
  if (typeof window === "undefined") return [];
  try {
    const stored = window.localStorage.getItem(storageKey);
    if (!stored) return [];
    const parsed: unknown = JSON.parse(stored);
    if (!Array.isArray(parsed)) return [];
    return parsed.filter(
      (report): report is AuditReport =>
        typeof report === "object" &&
        report !== null &&
        "id" in report &&
        typeof report.id === "string" &&
        "findings" in report &&
        Array.isArray(report.findings),
    );
  } catch {
    return [];
  }
}

function withoutRawUploads(report: AuditReport): AuditReport {
  return {
    ...report,
    files: report.files.map((file) => ({
      ...file,
      text: "",
      rows: [],
    })),
  };
}

export function listReports(includeDemo = true): AuditReport[] {
  const reports = readStoredReports().filter(
    (report) => report.id !== demoReport.id,
  );
  return [
    ...(includeDemo ? [demoReport] : []),
    ...reports.sort(
      (first, second) =>
        new Date(second.createdAt).getTime() -
        new Date(first.createdAt).getTime(),
    ),
  ];
}

export function getReport(id: string): AuditReport | undefined {
  if (id === demoReport.id) return demoReport;
  return readStoredReports().find((report) => report.id === id);
}

export function saveReport(report: AuditReport): void {
  if (typeof window === "undefined" || report.isDemo) return;
  const stored = readStoredReports().filter((item) => item.id !== report.id);
  stored.push(withoutRawUploads(report));
  try {
    window.localStorage.setItem(storageKey, JSON.stringify(stored));
  } catch {
    throw new Error(
      "This report could not be saved in this browser. Free up local storage and try again.",
    );
  }
}

export function deleteReport(id: string): void {
  if (typeof window === "undefined" || id === demoReport.id) return;
  const next = readStoredReports().filter((report) => report.id !== id);
  window.localStorage.setItem(storageKey, JSON.stringify(next));
}

export function getCategoryTotals(reports = listReports()): {
  category: FindingCategory;
  total: number | null;
  findingCount: number;
  hasUnpricedFindings: boolean;
}[] {
  const categories: FindingCategory[] = [
    "Unbilled work",
    "Scope creep",
    "Missing approval",
    "Missing evidence",
    "Over-delivery",
    "Missed commitment",
    "Billing mismatch",
  ];
  return categories.map((category) => {
    const findings = reports.flatMap((report) =>
      report.findings.filter((finding) => finding.category === category),
    );
    const hasUnpricedFindings = findings.some(
      (finding) => finding.estimatedValue === null,
    );
    return {
      category,
      total: hasUnpricedFindings
        ? null
        : findings.reduce(
            (sum, finding) => sum + (finding.estimatedValue ?? 0),
            0,
          ),
      findingCount: findings.length,
      hasUnpricedFindings,
    };
  });
}

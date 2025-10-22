/**
 * Tipos para el sistema de preflight de DOCX
 */

export type PreflightOptions = {
  keepComments: boolean;
  flattenSDT: boolean;
  acceptRevisions: boolean;
  customXmlPolicy: "keep" | "remove" | "auto";
};

export type PreflightReport = {
  fileParts: string[];
  fixes: string[];
  warnings: string[];
  styles: {
    total: number;
    deduped: number;
    defaultsFixed: number;
  };
  comments: {
    removed: number;
    kept: number;
  };
  numbering: {
    fixes: number;
  };
  contentTypes: {
    added: string[];
  };
  customXml: {
    action: "kept" | "removed";
    hasBindings: boolean;
  };
};

export type PreflightRun = {
  input: string;
  output: string;
  report: string;
  stepsDir?: string;
  analysisDir?: string;
  options?: Partial<PreflightOptions> & {
    verbose?: boolean;
    customXml?: "keep" | "remove" | "auto";
  };
};

export type DocxAnalysis = {
  parts: {
    document: boolean;
    styles: boolean;
    numbering: boolean;
    comments: boolean;
    rels: boolean;
    customXml: boolean;
  };
  comments: {
    entries: number;
    markersStart: number;
    markersEnd: number;
    refs: number;
  };
  styles: {
    total: number;
    defaults: number;
    normal: number;
    tableNormal: number;
  };
  numbering: {
    abstractNum: number;
    num: number;
    lvl: number;
  };
  contentTypes: {
    hasWdp: boolean;
    hasEmf: boolean;
  };
  rels: {
    total: number;
    images: number;
    missingMediaGuess: number;
  };
  altContent: {
    count: number;
  };
  revisions: {
    insertions: number;
    deletions: number;
  };
  sdt: {
    count: number;
    withBinding: number;
  };
  customXml: {
    items: number;
    hasBindings: boolean;
  };
}

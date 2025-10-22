import type { Logger } from "../util/log.js";
import type { PreflightReport } from "../core/types.js";

/**
 * Acepta todos los track changes:
 * - w:ins: sube su contenido (acepta inserción)
 * - w:del: elimina completamente (acepta eliminación)
 */
export function acceptAll(root: any, log: Logger, report: PreflightReport): any {
  let insAccepted = 0;
  let delRemoved = 0;

  const walk = (arr: any[]): any[] => {
    return arr.flatMap(node => {
      const k = Object.keys(node)[0];
      const el = node[k];

      if (k === "w:ins") {
        insAccepted++;
        // Acepta la inserción: sube el contenido
        const content = el?.["#text"] ?? [];
        return Array.isArray(content) ? content : [content];
      }

      if (k === "w:del") {
        delRemoved++;
        // Acepta la eliminación: borra todo
        return [];
      }

      // Recursión en hijos
      if (el && typeof el === "object" && el["#text"] && Array.isArray(el["#text"])) {
        el["#text"] = walk(el["#text"]);
      }

      return { [k]: el };
    });
  };

  const result = walk(root);

  if (insAccepted > 0 || delRemoved > 0) {
    report.fixes.push(`Revisions: ${insAccepted} inserciones aceptadas, ${delRemoved} eliminaciones aceptadas`);
    log.info({ insAccepted, delRemoved }, "Track changes aceptados");
  }

  return result;
}

/**
 * Verifica si un XML contiene track changes sin procesarlo completamente
 */
export function hasTrackChanges(xmlString: string): boolean {
  return xmlString.includes('w:ins>') || xmlString.includes('w:del>');
}

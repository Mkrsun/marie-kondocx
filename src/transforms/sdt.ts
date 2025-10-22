import type { Logger } from "../util/log.js";
import type { PreflightReport } from "../core/types.js";

/**
 * Aplana controles de contenido (SDT - Structured Document Tags)
 * Extrae el contenido de w:sdtContent y elimina la envoltura w:sdt
 */
export function flatten(root: any, log: Logger, report: PreflightReport): any {
  let count = 0;

  const walk = (arr: any[]): any[] => {
    return arr.flatMap(node => {
      const k = Object.keys(node)[0];
      const el = node[k];

      if (k === "w:sdt") {
        count++;
        // Buscar w:sdtContent en los hijos
        const children = (el?.["#text"] ?? []) as any[];
        const contentNode = children.find((n: any) => n["w:sdtContent"]);

        if (contentNode && contentNode["w:sdtContent"]) {
          const payload = contentNode["w:sdtContent"]["#text"] ?? [];
          log.debug("SDT aplanado, extrayendo contenido");
          return Array.isArray(payload) ? payload : [payload];
        }

        // SDT vacío o sin contenido
        log.debug("SDT vacío, eliminando");
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

  if (count > 0) {
    report.fixes.push(`SDT: ${count} controles aplanados`);
    log.info({ count }, "Content controls (SDT) aplanados");
  }

  return result;
}

/**
 * Verifica si un XML contiene SDT sin procesarlo completamente
 */
export function hasSDT(xmlString: string): boolean {
  return xmlString.includes('<w:sdt>') || xmlString.includes('<w:sdt ');
}

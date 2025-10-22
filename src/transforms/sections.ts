import type { Logger } from "../util/log.js";
import type { PreflightReport } from "../core/types.js";

/**
 * Verifica si un documento necesita sectPr final sin procesarlo completamente
 */
export function needsTrailingSectPr(xmlString: string): boolean {
  // Buscar si termina con w:sectPr antes de </w:body>
  const match = xmlString.match(/<\/w:body>/);
  if (!match) return false;
  
  const beforeBody = xmlString.substring(0, match.index);
  const lastSectPr = beforeBody.lastIndexOf('<w:sectPr');
  const lastP = beforeBody.lastIndexOf('<w:p ');
  const lastP2 = beforeBody.lastIndexOf('<w:p>');
  
  // Si no hay w:sectPr, o hay párrafos después del último sectPr, necesita uno final
  return lastSectPr === -1 || Math.max(lastP, lastP2) > lastSectPr;
}

/**
 * Asegura que haya un w:sectPr final en el documento
 * Word requiere que el último elemento de w:body sea w:sectPr
 */
export function ensureTrailingSectPr(root: any, log: Logger, report: PreflightReport): any {
  // Buscar w:document en el root array
  const docNode = root.find((n: any) => n["w:document"]);
  if (!docNode) {
    log.warn("No se encontró w:document");
    return root;
  }

  const doc = docNode["w:document"];

  // Después del procesamiento, w:document puede ser un array que contiene w:body directamente
  let bodyNode;
  if (Array.isArray(doc)) {
    // Estructura post-procesamiento: w:document es array [{ "w:body": [...] }]
    bodyNode = doc.find((n: any) => n["w:body"]);
  } else if (doc["#text"]) {
    // Estructura original: w:document tiene #text que contiene w:body
    bodyNode = doc["#text"].find((n: any) => n["w:body"]);
  } else if (doc["w:body"]) {
    // w:body directamente en w:document
    bodyNode = { "w:body": doc["w:body"] };
  }
  
  if (!bodyNode) {
    log.warn("No se encontró w:body");
    return root;
  }

  const body = bodyNode["w:body"];
  
  // Después del procesamiento, body puede ser array directamente o tener #text
  const kids = Array.isArray(body) ? body : (body["#text"] ?? []) as any[];

  // Verificar si el último elemento es w:sectPr
  const hasFinal = kids.length > 0 && Object.keys(kids[kids.length - 1])[0] === "w:sectPr";

  if (!hasFinal) {
    // Agregar w:sectPr básico al final
    kids.push({
      "w:sectPr": {
        "#text": [
          {
            "w:pgSz": {
              "@_w:w": "12240",
              "@_w:h": "15840"
            }
          },
          {
            "w:pgMar": {
              "@_w:top": "1440",
              "@_w:right": "1440",
              "@_w:bottom": "1440",
              "@_w:left": "1440",
              "@_w:header": "720",
              "@_w:footer": "720",
              "@_w:gutter": "0"
            }
          }
        ]
      }
    });

    body["#text"] = kids;
    report.fixes.push("Sections: w:sectPr final agregado");
    log.info("w:sectPr final agregado al documento");
  } else {
    log.debug("w:sectPr final ya existe");
  }

  return root;
}

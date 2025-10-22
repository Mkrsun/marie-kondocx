import type { Logger } from "../util/log.js";
import type { PreflightReport } from "../core/types.js";

/**
 * Verifica si un XML necesita desfragmentación de texto sin procesarlo completamente
 */
export function needsDefragmentation(xmlString: string): boolean {
  // Buscar patrones que indican fragmentación:
  // - Múltiples </w:t><w:t> consecutivos dentro del mismo run
  const fragmentPattern = /<\/w:t>\s*<w:t[^>]*>/g;
  return fragmentPattern.test(xmlString);
}

/**
 * Desfragmenta texto dentro de runs (w:r)
 * Une múltiples w:t consecutivos en un solo w:t para evitar que placeholders como {{nombre}}
 * queden fragmentados en múltiples nodos de texto
 */
export function defragmentRuns(root: any, log: Logger, report: PreflightReport): any {
  let runsFixed = 0;
  
  log.debug("Iniciando desfragmentación de texto");

  const walk = (node: any): any => {
    if (!node || typeof node !== "object") return node;

    // Si es un array, procesar cada elemento
    if (Array.isArray(node)) {
      return node.map(walk);
    }

    // Procesar el objeto actual
    const result = { ...node };
    
    // Buscar w:r que contenga múltiples w:t
    const keys = Object.keys(result);
    for (const k of keys) {
      const el = result[k];
      
      if (k === "w:r" && el && typeof el === "object" && el["#text"] && Array.isArray(el["#text"])) {
        const children = el["#text"] as any[];
        const textNodes = children.filter((n: any) => n && n["w:t"]);

        if (textNodes.length > 1) {
          runsFixed++;
          log.debug({ textNodes: textNodes.length }, "Desfragmentando w:t en run");
          
          // Fusionar todos los w:t en uno solo
          const mergedText = textNodes.map((t: any) => {
            const content = t["w:t"];
            // El texto puede estar en diferentes formatos según el parser
            if (typeof content === "string") return content;
            if (content && typeof content === "object" && content["#text"]) {
              const inner = content["#text"];
              return Array.isArray(inner) ? inner.join("") : String(inner);
            }
            return "";
          }).join("");

          // Mantener atributos del primer w:t (como xml:space)
          const firstText = textNodes[0]["w:t"];
          const attrs = (typeof firstText === "object" && firstText !== null && !Array.isArray(firstText))
            ? Object.keys(firstText).filter(k => k.startsWith("@_")).reduce((acc, k) => {
              acc[k] = firstText[k];
              return acc;
            }, {} as any)
            : {};

          // Preservar xml:space="preserve" si hay espacios
          if (mergedText !== mergedText.trim()) {
            attrs["@_xml:space"] = "preserve";
          }

          // Eliminar todos los w:t existentes y agregar uno nuevo fusionado
          const newChildren = children.filter((n: any) => !n || !n["w:t"]);
          newChildren.push({
            "w:t": {
              ...attrs,
              "#text": [mergedText]
            }
          });
          
          // Actualizar el elemento
          result[k] = {
            ...el,
            "#text": newChildren
          };

          log.debug({ merged: textNodes.length, mergedText: mergedText.substring(0, 50) }, "w:t desfragmentados en run");
        }
      }
      
      // Recursión en propiedades que son objetos o arrays
      if (el && typeof el === "object") {
        result[k] = walk(el);
      }
    }

    return result;
  };

  const result = walk(root);

  if (runsFixed > 0) {
    report.fixes.push(`Text: ${runsFixed} runs desfragmentados`);
    log.info({ runsFixed }, "Texto desfragmentado para preservar placeholders");
  }

  log.debug("Desfragmentación completada");
  return result;
}

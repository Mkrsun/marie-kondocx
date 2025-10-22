import type { Logger } from "../util/log.js";
import type { PreflightReport } from "../core/types.js";
import { Docx } from "../io/Docx.js";

/**
 * Verifica si el documento necesita procesamiento de CustomXML
 */
export async function needsCustomXmlProcessing(
  docx: Docx,
  policy: "keep" | "remove" | "auto"
): Promise<boolean> {
  if (policy === "keep") return false; // No necesita procesamiento
  if (policy === "remove") {
    // Solo necesita procesamiento si realmente hay customXml
    return docx.list("customXml/").length > 0;
  }
  
  // policy === "auto" - necesita procesamiento si hay customXml
  return docx.list("customXml/").length > 0;
}

/**
 * Aplica la política de CustomXML:
 * - keep: mantener customXml tal cual
 * - remove: eliminar completamente
 * - auto: detectar bindings y decidir automáticamente
 *
 * CustomXML almacena datos para vincular con controles de contenido (SDT).
 * Si no hay bindings (w:dataBinding), el customXml es residual y puede eliminarse.
 */
export async function applyPolicy(
  docx: Docx,
  policy: "keep" | "remove" | "auto",
  log: Logger,
  report: PreflightReport
): Promise<void> {
  const hasBindings = await detectBindings(docx);
  const shouldRemove = policy === "remove" || (policy === "auto" && !hasBindings);

  log.debug({ policy, hasBindings, shouldRemove }, "Política CustomXML");

  if (shouldRemove) {
    // Eliminar todos los archivos customXml
    const customXmlFiles = docx.list("customXml/");
    let removed = 0;

    log.debug({ customXmlFiles }, "Archivos CustomXML encontrados");

    for (const p of customXmlFiles) {
      log.debug({ file: p }, "Eliminando archivo CustomXML");
      docx.delete(p);
      removed++;
    }

    // Eliminar overrides de customXml en [Content_Types].xml
    const ct = (await docx.read("[Content_Types].xml")) ?? "";
    log.debug("Content_Types.xml antes de limpiar CustomXML");
    
    const cleaned = ct.replace(/<Override[^>]+PartName="\/customXml\/[^"]+"[^>]*\/>\s*/gi, "");

    if (cleaned !== ct) {
      docx.write("[Content_Types].xml", cleaned);
      log.debug("Overrides de customXml eliminados de Content_Types");
    } else {
      log.debug("No había overrides de CustomXML en Content_Types");
    }

    // Eliminar relaciones a customXml en _rels/.rels
    const rootRels = await docx.read("_rels/.rels");
    if (rootRels) {
      log.debug("Limpiando relaciones CustomXML de _rels/.rels");
      const cleanedRels = rootRels.replace(
        /<Relationship[^>]+Target="customXml\/[^"]+"[^>]*\/>\s*/gi,
        ""
      );
      if (cleanedRels !== rootRels) {
        docx.write("_rels/.rels", cleanedRels);
        log.debug("Relaciones customXml eliminadas de root .rels");
      } else {
        log.debug("No había relaciones CustomXML en root .rels");
      }
    }

    // Eliminar relaciones a customXml en word/_rels/document.xml.rels
    const docRels = await docx.read("word/_rels/document.xml.rels");
    if (docRels) {
      log.debug("Limpiando relaciones CustomXML de word/_rels/document.xml.rels");
      const cleanedDocRels = docRels.replace(
        /<Relationship[^>]+Target="[^"]*customXml\/[^"]+"[^>]*\/>\s*/gi,
        ""
      );
      if (cleanedDocRels !== docRels) {
        docx.write("word/_rels/document.xml.rels", cleanedDocRels);
        log.debug("Relaciones customXml eliminadas de document.xml.rels");
      } else {
        log.debug("No había relaciones CustomXML en document.xml.rels");
      }
    }

    report.fixes.push(`CustomXML: ${removed} archivos eliminados`);
    report.customXml = { action: "removed", hasBindings };
    log.info({ removed, hasBindings }, "CustomXML eliminado");
  } else {
    report.customXml = { action: "kept", hasBindings };
    log.info({ hasBindings, policy }, "CustomXML conservado");
  }
}

/**
 * Detecta si hay bindings de customXml en el documento
 * Busca w:dataBinding y w:storeItemID en document.xml, headers y footers
 */
async function detectBindings(docx: Docx): Promise<boolean> {
  const candidates = [
    "word/document.xml",
    ...docx.list("word/header"),
    ...docx.list("word/footer")
  ];

  for (const p of candidates) {
    const xml = await docx.read(p);
    if (!xml) continue;

    // Buscar indicadores de binding
    if (/w:dataBinding\b/i.test(xml) || /w:storeItemID\b/i.test(xml)) {
      return true;
    }
  }

  return false;
}

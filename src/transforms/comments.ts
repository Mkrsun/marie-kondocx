import type { Logger } from "../util/log.js";
import type { PreflightReport } from "../core/types.js";
import { Docx } from "../io/Docx.js";

/**
 * Verifica si hay comentarios en el documento sin procesarlo completamente
 */
export async function hasComments(docx: Docx): Promise<boolean> {
  // Verificar si existen archivos de comentarios
  const commentParts = [
    "word/comments.xml",
    "word/commentsExtended.xml",
    "word/commentsIds.xml"
  ];
  
  for (const part of commentParts) {
    if (docx.exists(part)) {
      return true;
    }
  }
  
  // Verificar si hay marcadores de comentarios en document.xml
  const docXml = await docx.read("word/document.xml");
  if (docXml && (
    docXml.includes('<w:commentRangeStart') ||
    docXml.includes('<w:commentRangeEnd') ||
    docXml.includes('<w:commentReference')
  )) {
    return true;
  }
  
  return false;
}

/**
 * Elimina todos los comentarios del documento:
 * - Borra archivos comments.xml y commentsExtended.xml
 * - Elimina marcadores de comentarios en document.xml y headers/footers
 */
export async function removeAll(docx: Docx, log: Logger): Promise<number> {
  let removed = 0;

  // Eliminar archivos de comentarios
  const commentParts = [
    "word/comments.xml",
    "word/commentsExtended.xml",
    "word/commentsIds.xml"
  ];

  for (const p of commentParts) {
    if (docx.exists(p)) {
      docx.delete(p);
      removed++;
      log.debug({ part: p }, "Archivo de comentarios eliminado");
    }
  }

  // Eliminar marcadores en todos los archivos de contenido
  const targets = [
    "word/document.xml",
    ...docx.list("word/header"),
    ...docx.list("word/footer")
  ];

  const patterns = [
    /<w:commentRangeStart[^>]*\/>/g,
    /<w:commentRangeEnd[^>]*\/>/g,
    /<w:commentReference[^>]*\/>/g
  ];

  for (const t of targets) {
    const xml = await docx.read(t);
    if (!xml) continue;

    let out = xml;
    for (const re of patterns) {
      out = out.replace(re, "");
    }

    if (out !== xml) {
      docx.write(t, out);
      log.debug({ part: t }, "Marcadores de comentarios eliminados");
      removed++;
    }
  }

  if (removed > 0) {
    log.info({ removed }, "Comentarios eliminados completamente");
  }

  return removed;
}

/**
 * Valida que los comentarios mantenidos estén consistentes
 */
export async function validate(docx: Docx, log: Logger, report: PreflightReport): Promise<void> {
  if (docx.exists("word/comments.xml")) {
    log.info("Comentarios conservados (keepComments=true)");

    // Validación básica: contar marcadores
    const doc = await docx.read("word/document.xml");
    if (doc) {
      const starts = (doc.match(/<w:commentRangeStart\b/g) ?? []).length;
      const ends = (doc.match(/<w:commentRangeEnd\b/g) ?? []).length;
      const refs = (doc.match(/<w:commentReference\b/g) ?? []).length;

      if (starts !== ends) {
        report.warnings.push(`Comentarios: marcadores desbalanceados (start:${starts}, end:${ends})`);
        log.warn({ starts, ends }, "Marcadores de comentarios desbalanceados");
      }

      if (refs === 0 && starts > 0) {
        report.warnings.push("Comentarios: hay rangos pero sin referencias");
        log.warn("Rangos de comentarios sin referencias");
      }
    }
  } else {
    log.debug("No hay comentarios en el documento");
  }
}

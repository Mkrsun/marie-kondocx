import type { Logger } from "../util/log.js";
import type { PreflightReport } from "../core/types.js";
import { Docx } from "../io/Docx.js";

/**
 * Verifica si hay problemas con las relaciones sin procesarlas completamente
 */
export async function hasRelationshipIssues(docx: Docx, log: Logger): Promise<boolean> {
  const relsPaths = [
    "word/_rels/document.xml.rels",
    ...docx.list("word/_rels/header").filter(p => p.endsWith(".rels")),
    ...docx.list("word/_rels/footer").filter(p => p.endsWith(".rels"))
  ];

  for (const relPath of relsPaths) {
    const xml = await docx.read(relPath);
    if (!xml) continue;

    const rels = [...xml.matchAll(/<Relationship\b([^>]+)\/>/g)];
    
    for (const match of rels) {
      const attrs = match[1];
      const targetMatch = attrs.match(/Target="([^"]+)"/);
      const target = targetMatch?.[1] ?? "";

      if (!target || /^https?:\/\//i.test(target)) continue;

      // Normalizar path
      const abs = normalizeTarget(relPath, target);
      
      // Si hay un archivo que no existe, necesita corrección
      if (!docx.exists(abs)) {
        return true;
      }
    }
  }
  
  return false;
}

/**
 * Valida y corrige relaciones (.rels):
 * - Verifica que los targets existan
 * - Elimina relaciones huérfanas (que apuntan a archivos inexistentes)
 * - Reporta relaciones externas
 */
export async function validateAndFix(docx: Docx, log: Logger, report: PreflightReport): Promise<void> {
  const relsPaths = [
    "word/_rels/document.xml.rels",
    ...docx.list("word/_rels/header").filter(p => p.endsWith(".rels")),
    ...docx.list("word/_rels/footer").filter(p => p.endsWith(".rels"))
  ];

  for (const relPath of relsPaths) {
    const xml = await docx.read(relPath);
    if (!xml) continue;

    const rels = [...xml.matchAll(/<Relationship\b([^>]+)\/>/g)];
    let fixed = xml;
    let removedCount = 0;

    for (const match of rels) {
      const attrs = match[1];
      const idMatch = attrs.match(/Id="([^"]+)"/);
      const targetMatch = attrs.match(/Target="([^"]+)"/);
      const typeMatch = attrs.match(/Type="([^"]+)"/);

      const id = idMatch?.[1] ?? "";
      const target = targetMatch?.[1] ?? "";
      const type = typeMatch?.[1] ?? "";

      if (!target) continue;

      // Skip relaciones externas
      if (/^https?:\/\//i.test(target)) {
        log.debug({ id, target }, "Relación externa (skip)");
        continue;
      }

      // Normalizar path
      const abs = normalizeTarget(relPath, target);

      // Verificar existencia
      if (!docx.exists(abs)) {
        // Si es media y no existe, es problema
        if (abs.startsWith("word/media/")) {
          report.warnings.push(`Rels: media faltante ${abs} (Id=${id})`);
          log.warn({ relPath, id, target: abs }, "Relación apunta a media inexistente");

          // Eliminar relación huérfana
          fixed = fixed.replace(match[0], "");
          removedCount++;
        } else if (type.includes("comments") || type.includes("endnotes") || type.includes("footnotes")) {
          // Partes opcionales que podrían no existir - eliminar relación
          log.debug({ relPath, id, target: abs }, "Relación a parte opcional inexistente, eliminando");
          fixed = fixed.replace(match[0], "");
          removedCount++;
        } else {
          log.debug({ relPath, id, target: abs, type }, "Relación a archivo inexistente");
        }
      }
    }

    if (fixed !== xml) {
      docx.write(relPath, fixed);
      report.fixes.push(`Rels: ${removedCount} relaciones huérfanas eliminadas en ${relPath}`);
      log.info({ relPath, removed: removedCount }, "Relaciones huérfanas eliminadas");
    }
  }
}

/**
 * Normaliza un target path relativo a absoluto dentro del ZIP
 */
function normalizeTarget(relPath: string, target: string): string {
  // Si es URL externa, devolver tal cual
  if (/^https?:\/\//i.test(target)) return target;

  // Si es absoluto (empieza con /), quitar el /
  if (target.startsWith("/")) return target.slice(1);

  // Es relativo, resolverlo desde la ubicación del .rels
  // Ej: word/_rels/document.xml.rels + media/image1.png = word/media/image1.png
  const base = relPath.replace(/_rels\/[^/]+\.rels$/, "");
  return base + target;
}

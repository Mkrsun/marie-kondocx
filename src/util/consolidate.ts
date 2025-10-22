import { readFile, writeFile } from "node:fs/promises";
import { readdirSync, existsSync } from "node:fs";
import { join, basename } from "node:path";
import type { PreflightReport, DocxAnalysis } from "../core/types.js";

export type ConsolidatedEntry = {
  filename: string;
  // Análisis ANTES
  before: {
    altContent: number;
    comments: number;
    styles: number;
    sdt: number;
    customXml: number;
    revisions: number;
    rels: number;
  };
  // Análisis DESPUÉS
  after: {
    altContent: number;
    comments: number;
    styles: number;
    sdt: number;
    customXml: number;
    revisions: number;
    rels: number;
  };
  // Reporte de cambios
  fixes: {
    total: number;
    altContentProcessed: number;
    commentsRemoved: number;
    stylesDeduped: number;
    numberingFixes: number;
    contentTypesAdded: number;
    customXmlAction: string;
  };
  // Metadatos
  processingTime?: number;
  success: boolean;
  error?: string;
};

/**
 * Consolida todos los reportes JSON individuales en un archivo CSV
 */
export async function generateConsolidatedReport(
  logsDir: string = "logs",
  outputFile: string = "consolidated-report.csv"
): Promise<void> {
  console.log(`\n🔍 Generando reporte consolidado...`);
  
  if (!existsSync(logsDir)) {
    throw new Error(`La carpeta ${logsDir} no existe`);
  }

  // Buscar todos los archivos de reporte
  const files = readdirSync(logsDir).filter(file => 
    file.endsWith('-report.json') && !file.startsWith('consolidated')
  );

  if (files.length === 0) {
    throw new Error(`No se encontraron archivos de reporte en ${logsDir}`);
  }

  console.log(`📊 Procesando ${files.length} reportes...`);

  const consolidatedData: ConsolidatedEntry[] = [];

  for (const reportFile of files) {
    try {
      const reportPath = join(logsDir, reportFile);
      const analysisBaseName = reportFile.replace('-report.json', '');
      const beforeFile = join("analysis", `${analysisBaseName}.before.json`);
      const afterFile = join("analysis", `${analysisBaseName}.after.json`);

      // Leer reporte
      const reportContent = await readFile(reportPath, 'utf8');
      const report: PreflightReport = JSON.parse(reportContent);

      // Leer análisis ANTES (si existe)
      let beforeAnalysis: DocxAnalysis | null = null;
      if (existsSync(beforeFile)) {
        const beforeContent = await readFile(beforeFile, 'utf8');
        beforeAnalysis = JSON.parse(beforeContent);
      }

      // Leer análisis DESPUÉS (si existe)
      let afterAnalysis: DocxAnalysis | null = null;
      if (existsSync(afterFile)) {
        const afterContent = await readFile(afterFile, 'utf8');
        afterAnalysis = JSON.parse(afterContent);
      }

      // Extraer número de elementos procesados de AlternateContent
      const altContentMatch = report.fixes.find(fix => 
        fix.includes('AlternateContent:') && fix.includes('procesados')
      );
      const altContentProcessed = altContentMatch 
        ? parseInt(altContentMatch.match(/(\d+)\/\d+ procesados/)?.[1] || '0')
        : 0;

      const entry: ConsolidatedEntry = {
        filename: analysisBaseName,
        before: {
          altContent: beforeAnalysis?.altContent.count || 0,
          comments: beforeAnalysis?.comments.entries || 0,
          styles: beforeAnalysis?.styles.total || 0,
          sdt: beforeAnalysis?.sdt.count || 0,
          customXml: beforeAnalysis?.customXml.items || 0,
          revisions: (beforeAnalysis?.revisions.insertions || 0) + (beforeAnalysis?.revisions.deletions || 0),
          rels: beforeAnalysis?.rels.total || 0,
        },
        after: {
          altContent: afterAnalysis?.altContent.count || 0,
          comments: afterAnalysis?.comments.entries || 0,
          styles: afterAnalysis?.styles.total || 0,
          sdt: afterAnalysis?.sdt.count || 0,
          customXml: afterAnalysis?.customXml.items || 0,
          revisions: (afterAnalysis?.revisions.insertions || 0) + (afterAnalysis?.revisions.deletions || 0),
          rels: afterAnalysis?.rels.total || 0,
        },
        fixes: {
          total: report.fixes.length,
          altContentProcessed,
          commentsRemoved: report.comments.removed,
          stylesDeduped: report.styles.deduped,
          numberingFixes: report.numbering.fixes,
          contentTypesAdded: report.contentTypes.added.length,
          customXmlAction: report.customXml.action,
        },
        success: true,
      };

      consolidatedData.push(entry);
      console.log(`  ✓ ${analysisBaseName}`);

    } catch (error) {
      const filename = reportFile.replace('-report.json', '');
      console.warn(`  ⚠️  Error procesando ${filename}: ${error}`);
      
      consolidatedData.push({
        filename,
        before: { altContent: 0, comments: 0, styles: 0, sdt: 0, customXml: 0, revisions: 0, rels: 0 },
        after: { altContent: 0, comments: 0, styles: 0, sdt: 0, customXml: 0, revisions: 0, rels: 0 },
        fixes: { total: 0, altContentProcessed: 0, commentsRemoved: 0, stylesDeduped: 0, numberingFixes: 0, contentTypesAdded: 0, customXmlAction: "error" },
        success: false,
        error: String(error),
      });
    }
  }

  // Generar CSV
  const csvHeader = [
    'filename',
    'success',
    'error',
    // ANTES
    'before_altcontent',
    'before_comments', 
    'before_styles',
    'before_sdt',
    'before_customxml',
    'before_revisions',
    'before_rels',
    // DESPUÉS
    'after_altcontent',
    'after_comments',
    'after_styles', 
    'after_sdt',
    'after_customxml',
    'after_revisions',
    'after_rels',
    // CAMBIOS
    'fixes_total',
    'altcontent_processed',
    'comments_removed',
    'styles_deduped',
    'numbering_fixes',
    'contenttypes_added',
    'customxml_action',
    // COMPLEJIDAD
    'complexity_score',
    'processing_impact'
  ].join(',');

  const csvRows = consolidatedData.map(entry => {
    // Calcular score de complejidad (elementos problemáticos ANTES)
    const complexityScore = entry.before.altContent * 3 +  // AlternateContent es lo más complejo
                           entry.before.comments * 2 +     // Comentarios moderadamente complejos
                           entry.before.sdt * 2 +          // SDTs moderadamente complejos
                           entry.before.customXml * 1 +    // CustomXML menos complejo
                           (entry.before.revisions > 0 ? 2 : 0); // Revisiones moderadas

    // Calcular impacto del procesamiento
    const processingImpact = entry.fixes.altContentProcessed +
                            entry.fixes.commentsRemoved +
                            entry.fixes.stylesDeduped +
                            entry.fixes.numberingFixes +
                            (entry.fixes.customXmlAction === 'removed' ? 1 : 0);

    return [
      entry.filename,
      entry.success,
      entry.error || '',
      // ANTES
      entry.before.altContent,
      entry.before.comments,
      entry.before.styles,
      entry.before.sdt,
      entry.before.customXml,
      entry.before.revisions,
      entry.before.rels,
      // DESPUÉS
      entry.after.altContent,
      entry.after.comments,
      entry.after.styles,
      entry.after.sdt,
      entry.after.customXml,
      entry.after.revisions,
      entry.after.rels,
      // CAMBIOS
      entry.fixes.total,
      entry.fixes.altContentProcessed,
      entry.fixes.commentsRemoved,
      entry.fixes.stylesDeduped,
      entry.fixes.numberingFixes,
      entry.fixes.contentTypesAdded,
      entry.fixes.customXmlAction,
      // COMPLEJIDAD
      complexityScore,
      processingImpact
    ].join(',');
  });

  const csvContent = [csvHeader, ...csvRows].join('\n');

  // Escribir CSV
  await writeFile(outputFile, csvContent, 'utf8');

  // Estadísticas de resumen
  const totalFiles = consolidatedData.length;
  const successfulFiles = consolidatedData.filter(e => e.success).length;
  const failedFiles = totalFiles - successfulFiles;
  const totalAltContentBefore = consolidatedData.reduce((sum, e) => sum + e.before.altContent, 0);
  const totalAltContentProcessed = consolidatedData.reduce((sum, e) => sum + e.fixes.altContentProcessed, 0);
  const avgComplexity = consolidatedData.reduce((sum, e) => {
    const score = e.before.altContent * 3 + e.before.comments * 2 + e.before.sdt * 2 + e.before.customXml * 1;
    return sum + score;
  }, 0) / totalFiles;

  console.log(`\n📋 Reporte consolidado generado: ${outputFile}`);
  console.log(`\n📊 ESTADÍSTICAS GENERALES:`);
  console.log(`   Total archivos procesados: ${totalFiles}`);
  console.log(`   Exitosos: ${successfulFiles}`);
  console.log(`   Con errores: ${failedFiles}`);
  console.log(`   AlternateContent total (antes): ${totalAltContentBefore}`);
  console.log(`   AlternateContent procesado: ${totalAltContentProcessed}`);
  console.log(`   Complejidad promedio: ${avgComplexity.toFixed(1)}`);
  console.log(`\n💡 El archivo CSV puede importarse directamente a Excel para análisis detallado.`);
}
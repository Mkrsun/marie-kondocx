import path from "node:path";
import { mkdir, writeFile } from "node:fs/promises";
import { Docx } from "./io/Docx.js";
import { PreflightService } from "./core/PreflightService.js";
import { createLogger, createFileLogger } from "./util/log.js";
import type { PreflightRun } from "./core/types.js";
import { analyzeDocx, formatAnalysis } from "./core/analyzer.js";

/**
 * Función principal que ejecuta el preflight con análisis incremental
 */
export async function preflightFile(run: PreflightRun): Promise<void> {
  // Crear directorios de salida
  const stepsDir = run.stepsDir;
  const analysisDir = run.analysisDir;
  const logsDir = "logs";

  if (stepsDir) {
    await mkdir(stepsDir, { recursive: true });
  }
  if (analysisDir) {
    await mkdir(analysisDir, { recursive: true });
  }
  await mkdir(logsDir, { recursive: true });

  const base = path.basename(run.input).replace(/\.docx$/i, "");

  // Paths para archivos generados (solo si se especifican las carpetas)
  const stepDocx = stepsDir ? (suffix: string) => path.join(stepsDir, `${base}.${suffix}.docx`) : null;
  const analysisJson = analysisDir ? (suffix: string) => path.join(analysisDir, `${base}.${suffix}.json`) : null;
  const logFile = path.join(logsDir, `${base}.log`);
  const reportFile = path.join(logsDir, `${base}-report.json`);

  // Crear logger que escribe a consola y archivo
  const log = createFileLogger(run.options?.verbose ? "debug" : "info", logFile);

  log.info({ input: run.input }, "Abriendo documento");

  // Abrir documento
  const docx = await Docx.open(run.input);

  // Análisis inicial
  log.info("Analizando estructura inicial del documento");
  const before = await analyzeDocx(docx);
  if (analysisJson) {
    await writeFile(analysisJson("before"), JSON.stringify(before, null, 2), "utf8");
  }

  // Mostrar análisis en consola
  console.log(formatAnalysis(before, "Análisis ANTES del preflight"));

  // Guardar copia original como paso 00
  if (stepDocx) {
    await docx.saveAs(stepDocx("00-original"));
    log.info({ path: stepDocx("00-original") }, "Copia original guardada");
  }

  // Crear servicio de preflight con hooks para snapshots
  const svc = new PreflightService(
    log,
    {
      keepComments: run.options?.keepComments ?? false,
      flattenSDT: run.options?.flattenSDT ?? true,
      acceptRevisions: run.options?.acceptRevisions ?? true,
      customXmlPolicy: run.options?.customXml ?? "auto"
    },
    {
      snapshot: async (suffix) => {
        // Guardar DOCX del paso (si se habilitó)
        if (stepDocx) {
          const docxPath = stepDocx(suffix);
          await docx.saveAs(docxPath);
          log.info({ step: suffix, docx: docxPath }, "Checkpoint DOCX guardado");
        }

        // Analizar estado después del paso (si se habilitó)
        if (analysisJson) {
          const snap = await analyzeDocx(docx);
          const jsonPath = analysisJson(`after-step-${suffix}`);
          await writeFile(jsonPath, JSON.stringify(snap, null, 2), "utf8");
          log.info({ step: suffix, analysis: jsonPath }, "Checkpoint análisis guardado");
        }

        if (!stepDocx && !analysisJson) {
          log.debug({ step: suffix }, "Checkpoint sin archivos - procesamiento directo");
        }
      }
    }
  );

  // Ejecutar preflight
  log.info("Iniciando proceso de preflight");
  const report = await svc.run(docx);

  // Guardar paso final (si se habilitó)
  if (stepDocx) {
    await docx.saveAs(stepDocx("99-final"));
    log.info({ path: stepDocx("99-final") }, "Documento final guardado");
  }

  // Análisis final
  log.info("Analizando estructura final del documento");
  const after = await analyzeDocx(docx);
  if (analysisJson) {
    await writeFile(analysisJson("after"), JSON.stringify(after, null, 2), "utf8");
  }

  // Mostrar análisis final en consola
  console.log(formatAnalysis(after, "Análisis DESPUÉS del preflight"));

  // Guardar output solicitado
  await docx.saveAs(run.output);
  await writeFile(run.report, JSON.stringify(report, null, 2), "utf8");
  
  // Guardar reporte también en el archivo dedicado por documento
  await writeFile(reportFile, JSON.stringify(report, null, 2), "utf8");

  // Resumen de cambios
  console.log("\n=== RESUMEN DE CAMBIOS ===\n");
  console.log(`Fixes aplicados: ${report.fixes.length}`);
  report.fixes.forEach(fix => console.log(`  - ${fix}`));

  if (report.warnings.length > 0) {
    console.log(`\nWarnings: ${report.warnings.length}`);
    report.warnings.forEach(warn => console.log(`  - ${warn}`));
  }

  console.log(`\nEstilos: ${report.styles.total} (${report.styles.deduped} deduplicados, ${report.styles.defaultsFixed} defaults corregidos)`);
  console.log(`Comentarios: ${report.comments.removed} eliminados, ${report.comments.kept} conservados`);
  console.log(`Numeración: ${report.numbering.fixes} fixes`);
  console.log(`Content Types: ${report.contentTypes.added.length} agregados [${report.contentTypes.added.join(", ")}]`);
  console.log(`CustomXML: ${report.customXml.action} (bindings: ${report.customXml.hasBindings})`);

  log.info({
    output: run.output,
    report: run.report,
    logFile,
    reportFile,
    stepsDir,
    analysisDir
  }, "Preflight terminado exitosamente");

  console.log(`\n📋 Archivos generados:`);
  console.log(`  - Log detallado: ${logFile}`);
  console.log(`  - Reporte JSON: ${reportFile}`);
}

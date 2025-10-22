#!/usr/bin/env node
import yargs from "yargs";
import { hideBin } from "yargs/helpers";
import { preflightFile } from "./index.js";
import { generateConsolidatedReport } from "./util/consolidate.js";
import { readdirSync, existsSync } from "fs";
import { join, basename, extname } from "path";

yargs(hideBin(process.argv))
  .command(
    "run",
    "Ejecuta preflight y limpieza de un documento DOCX",
    (y) =>
      y
        .option("input", {
          alias: "i",
          type: "string",
          demandOption: true,
          describe: "Archivo .docx de entrada"
        })
        .option("output", {
          alias: "o",
          type: "string",
          demandOption: true,
          describe: "Archivo .docx de salida (limpio)"
        })
        .option("report", {
          alias: "r",
          type: "string",
          default: "report.json",
          describe: "Archivo de reporte JSON"
        })
        .option("stepsDir", {
          type: "string",
          describe: "Directorio para archivos DOCX intermedios por paso (opcional)"
        })
        .option("analysisDir", {
          type: "string", 
          describe: "Directorio para análisis JSON por paso (opcional)"
        })
        .option("keepComments", {
          type: "boolean",
          default: false,
          describe: "No eliminar comentarios"
        })
        .option("flattenSDT", {
          type: "boolean",
          default: true,
          describe: "Aplanar controles de contenido (SDT)"
        })
        .option("acceptRevisions", {
          type: "boolean",
          default: true,
          describe: "Aceptar todos los track changes"
        })
        .option("customXml", {
          choices: ["keep", "remove", "auto"] as const,
          default: "auto",
          describe: "Política de CustomXML: keep/remove/auto"
        })
        .option("verbose", {
          alias: "v",
          type: "boolean",
          default: false,
          describe: "Logging detallado"
        }),
    async (argv) => {
      try {
        await preflightFile({
          input: String(argv.input),
          output: String(argv.output),
          report: String(argv.report),
          ...(argv.stepsDir && { stepsDir: String(argv.stepsDir) }),
          ...(argv.analysisDir && { analysisDir: String(argv.analysisDir) }),
          options: {
            keepComments: Boolean(argv.keepComments),
            flattenSDT: Boolean(argv.flattenSDT),
            acceptRevisions: Boolean(argv.acceptRevisions),
            customXml: argv.customXml as "keep" | "remove" | "auto",
            verbose: Boolean(argv.verbose)
          }
        });
        process.exit(0);
      } catch (error) {
        console.error("Error durante el preflight:", error);
        process.exit(1);
      }
    }
  )
  .command(
    "process-templates",
    "Procesa todos los archivos .docx en la carpeta templates/",
    (y) =>
      y
        .option("templatesDir", {
          type: "string",
          default: "templates",
          describe: "Directorio con los templates a procesar"
        })
        .option("resultsDir", {
          type: "string",
          default: "results",
          describe: "Directorio donde se guardarán los archivos limpios"
        })
        .option("report", {
          alias: "r",
          type: "string",
          default: "report.json",
          describe: "Archivo de reporte JSON"
        })
        .option("stepsDir", {
          type: "string",
          describe: "Directorio para archivos DOCX intermedios por paso (opcional)"
        })
        .option("analysisDir", {
          type: "string",
          describe: "Directorio para análisis JSON por paso (opcional)"
        })
        .option("keepComments", {
          type: "boolean",
          default: false,
          describe: "No eliminar comentarios"
        })
        .option("flattenSDT", {
          type: "boolean",
          default: true,
          describe: "Aplanar controles de contenido (SDT)"
        })
        .option("acceptRevisions", {
          type: "boolean",
          default: true,
          describe: "Aceptar todos los track changes"
        })
        .option("customXml", {
          choices: ["keep", "remove", "auto"] as const,
          default: "auto",
          describe: "Política de CustomXML: keep/remove/auto"
        })
        .option("verbose", {
          alias: "v",
          type: "boolean",
          default: false,
          describe: "Logging detallado"
        })
        .option("consolidatedReport", {
          type: "string",
          default: "consolidated-report.csv",
          describe: "Archivo CSV con reporte consolidado de todos los templates"
        }),
    async (argv) => {
      try {
        const templatesDir = String(argv.templatesDir);
        const resultsDir = String(argv.resultsDir);

        // Verificar que la carpeta templates existe
        if (!existsSync(templatesDir)) {
          console.error(`Error: La carpeta ${templatesDir} no existe.`);
          process.exit(1);
        }

        // Leer todos los archivos .docx en la carpeta templates
        const files = readdirSync(templatesDir).filter(
          (file) => extname(file).toLowerCase() === ".docx"
        );

        if (files.length === 0) {
          console.log(`No se encontraron archivos .docx en ${templatesDir}`);
          process.exit(0);
        }

        console.log(`Encontrados ${files.length} archivos para procesar:`);
        files.forEach((file) => console.log(`  - ${file}`));
        console.log("");

        // Procesar cada archivo
        let processed = 0;
        let failed = 0;

        for (const file of files) {
          const inputPath = join(templatesDir, file);
          const outputPath = join(resultsDir, file);

          console.log(`\nProcesando [${processed + 1}/${files.length}]: ${file}`);

          try {
            await preflightFile({
              input: inputPath,
              output: outputPath,
              report: String(argv.report),
              ...(argv.stepsDir && { stepsDir: String(argv.stepsDir) }),
              ...(argv.analysisDir && { analysisDir: String(argv.analysisDir) }),
              options: {
                keepComments: Boolean(argv.keepComments),
                flattenSDT: Boolean(argv.flattenSDT),
                acceptRevisions: Boolean(argv.acceptRevisions),
                customXml: argv.customXml as "keep" | "remove" | "auto",
                verbose: Boolean(argv.verbose)
              }
            });

            console.log(`✓ Completado: ${file} -> ${outputPath}`);
            processed++;
          } catch (error) {
            console.error(`✗ Error procesando ${file}:`, error);
            failed++;
          }
        }

        console.log(`\n${"=".repeat(60)}`);
        console.log(`Resumen:`);
        console.log(`  Total archivos: ${files.length}`);
        console.log(`  Procesados exitosamente: ${processed}`);
        console.log(`  Fallidos: ${failed}`);
        console.log(`${"=".repeat(60)}\n`);

        // Generar reporte consolidado
        try {
          await generateConsolidatedReport("logs", String(argv.consolidatedReport));
        } catch (error) {
          console.warn(`⚠️  No se pudo generar el reporte consolidado: ${error}`);
        }

        process.exit(failed > 0 ? 1 : 0);
      } catch (error) {
        console.error("Error durante el procesamiento de templates:", error);
        process.exit(1);
      }
    }
  )
  .demandCommand(1, "Debes especificar un comando")
  .strict()
  .help()
  .alias("help", "h")
  .version("1.0.0")
  .alias("version", "V")
  .epilogue("Para más información: https://github.com/tu-usuario/marie-kondocx")
  .parse();

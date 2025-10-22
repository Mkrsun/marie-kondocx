import type { Logger } from "../util/log.js";
import type { PreflightOptions, PreflightReport } from "./types.js";
import { Docx } from "../io/Docx.js";
import { parse, build, ensureWNamespace } from "../io/Xml.js";
import * as Alt from "../transforms/altcontent.js";
import * as Rev from "../transforms/revisions.js";
import * as SDT from "../transforms/sdt.js";
import * as Txt from "../transforms/text.js";
import * as Sty from "../transforms/styles.js";
import * as Num from "../transforms/numbering.js";
import * as Cmt from "../transforms/comments.js";
import * as Rels from "../transforms/rels.js";
import * as CT from "../transforms/contentTypes.js";
import * as Sec from "../transforms/sections.js";
import * as CXml from "../transforms/customxml.js";

type Hooks = {
  snapshot: (suffix: string) => Promise<void>;
};

/**
 * Servicio principal que orquesta todos los pasos de preflight
 * Cada paso genera un checkpoint (DOCX + análisis)
 */
export class PreflightService {
  constructor(
    private log: Logger,
    private opts: PreflightOptions,
    private hooks: Hooks
  ) {}

  async run(docx: Docx): Promise<PreflightReport> {
    const report: PreflightReport = {
      fileParts: docx.list(),
      fixes: [],
      warnings: [],
      styles: { total: 0, deduped: 0, defaultsFixed: 0 },
      comments: { removed: 0, kept: 0 },
      numbering: { fixes: 0 },
      contentTypes: { added: [] },
      customXml: { action: "kept", hasBindings: false }
    };

    this.log.info("Iniciando preflight con checkpoints por paso");

    // 01 - AlternateContent
    await this.step("01-altcontent", async () => {
      const xml = (await docx.read("word/document.xml")) ?? "<w:document/>";
      
      // Optimización: Solo procesar si realmente hay AlternateContent
      if (Alt.hasAlternateContent(xml)) {
        this.log.info("🚀 INICIANDO procesamiento de AlternateContent");
        
        // LOG DEL XML CRUDO para debug
        const altContentMatches = (xml.match(/mc:AlternateContent/g) || []).length;
        this.log.info({ 
          xmlLength: xml.length, 
          altContentMatches,
          xmlPreview: xml.substring(0, 500) + "..."
        }, "📝 XML CRUDO contiene AlternateContent");
        
        this.log.info("🔄 Llamando a resolveFallbacks...");
        const processedXml = Alt.resolveFallbacks(xml, this.log, report);
        this.log.info("✅ resolveFallbacks completado, escribiendo XML...");
        docx.write("word/document.xml", processedXml);
        this.log.info("AlternateContent procesado - XML modificado");
      } else {
        this.log.debug("No hay AlternateContent - XML preservado sin cambios");
        // No escribir nada, mantener XML original intacto
      }
    });

    // 02 - Revisions (Track Changes)
    if (this.opts.acceptRevisions) {
      await this.step("02-revisions", async () => {
        const xml = (await docx.read("word/document.xml"))!;
        
        // Optimización: Solo procesar si realmente hay track changes
        if (Rev.hasTrackChanges(xml)) {
          let doc = parse(xml);
          doc = Rev.acceptAll(doc, this.log, report);
          docx.write("word/document.xml", build(doc));
          this.log.info("Track changes procesados - XML modificado");
        } else {
          this.log.debug("No hay track changes - XML preservado sin cambios");
          // No escribir nada, mantener XML original intacto
        }
      });
    } else {
      this.log.info("Paso 02-revisions omitido (acceptRevisions=false)");
    }

    // 03 - SDT (Content Controls)
    if (this.opts.flattenSDT) {
      await this.step("03-sdt", async () => {
        const xml = (await docx.read("word/document.xml"))!;
        
        // Optimización: Solo procesar si realmente hay SDT
        if (SDT.hasSDT(xml)) {
          let doc = parse(xml);
          doc = SDT.flatten(doc, this.log, report);
          docx.write("word/document.xml", build(doc));
          this.log.info("SDT procesados - XML modificado");
        } else {
          this.log.debug("No hay SDT - XML preservado sin cambios");
          // No escribir nada, mantener XML original intacto
        }
      });
    } else {
      this.log.info("Paso 03-sdt omitido (flattenSDT=false)");
    }

    // 04 - Text defragmentation
    await this.step("04-text", async () => {
      const xml = (await docx.read("word/document.xml"))!;
      
      // Optimización: Solo procesar si realmente necesita desfragmentación
      if (Txt.needsDefragmentation(xml)) {
        let doc = parse(xml);
        doc = Txt.defragmentRuns(doc, this.log, report);
        docx.write("word/document.xml", build(doc));
        this.log.info("Texto desfragmentado - XML modificado");
      } else {
        this.log.debug("No necesita desfragmentación - XML preservado sin cambios");
        // No escribir nada, mantener XML original intacto
      }
    });

    // 05 - Styles
    await this.step("05-styles", async () => {
      let stylesXml = await docx.read("word/styles.xml");
      if (!stylesXml) {
        stylesXml = '<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>';
        docx.write("word/styles.xml", stylesXml);
        this.log.info("Styles.xml creado");
      } else if (Sty.needsSanitization(stylesXml)) {
        let styles = parse(stylesXml);
        ensureWNamespace(styles);
        styles = Sty.sanitize(styles, this.log, report);
        docx.write("word/styles.xml", build(styles));
        this.log.info("Styles procesados - XML modificado");
      } else {
        this.log.debug("Styles no necesitan saneamiento - XML preservado sin cambios");
        // No escribir nada, mantener XML original intacto
      }
    });

    // 06 - Numbering
    await this.step("06-numbering", async () => {
      let numberingXml = (await docx.read("word/numbering.xml")) ??
        '<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>';
      
      if (Num.needsCleaning(numberingXml)) {
        let numbering = parse(numberingXml);
        ensureWNamespace(numbering);
        numbering = Num.cleanAndNormalize(numbering, this.log, report);
        docx.write("word/numbering.xml", build(numbering));
        this.log.info("Numbering procesado - XML modificado");
      } else {
        this.log.debug("Numbering no necesita limpieza - XML preservado sin cambios");
        // No escribir nada, mantener XML original intacto
      }
    });

    // 07 - Comments
    if (!this.opts.keepComments) {
      await this.step("07-comments", async () => {
        if (await Cmt.hasComments(docx)) {
          const removed = await Cmt.removeAll(docx, this.log);
          report.comments.removed = removed;
          this.log.info("Comentarios eliminados - archivos modificados");
        } else {
          this.log.debug("No hay comentarios - archivos preservados sin cambios");
          report.comments.removed = 0;
        }
      });
    } else {
      await this.step("07-comments-keep", async () => {
        await Cmt.validate(docx, this.log, report);
        report.comments.kept = 1;
      });
    }

    // 08 - Rels
    await this.step("08-rels", async () => {
      if (await Rels.hasRelationshipIssues(docx, this.log)) {
        await Rels.validateAndFix(docx, this.log, report);
        this.log.info("Relaciones corregidas - archivos modificados");
      } else {
        this.log.debug("Relaciones válidas - archivos preservados sin cambios");
      }
    });

    // 09 - Content Types
    await this.step("09-contenttypes", async () => {
      if (await CT.needsContentTypes(docx)) {
        report.contentTypes.added = await CT.ensure(docx, this.log);
        this.log.info("Content types actualizados - archivo modificado");
      } else {
        this.log.debug("Content types completos - archivo preservado sin cambios");
        report.contentTypes.added = [];
      }
    });

    // 10 - Sections
    await this.step("10-sections", async () => {
      const xml = (await docx.read("word/document.xml"))!;
      
      if (Sec.needsTrailingSectPr(xml)) {
        let doc = parse(xml);
        doc = Sec.ensureTrailingSectPr(doc, this.log, report);
        docx.write("word/document.xml", build(doc));
        this.log.info("SectPr añadido - XML modificado");
      } else {
        this.log.debug("SectPr final ya existe - XML preservado sin cambios");
      }
    });

    // 11 - CustomXML
    await this.step("11-customxml", async () => {
      if (await CXml.needsCustomXmlProcessing(docx, this.opts.customXmlPolicy)) {
        await CXml.applyPolicy(docx, this.opts.customXmlPolicy, this.log, report);
        this.log.info("CustomXML procesado - archivos modificados");
      } else {
        this.log.debug("CustomXML no necesita procesamiento - archivos preservados sin cambios");
        report.customXml = { action: "kept", hasBindings: false };
      }
    });

    this.log.info({ fixes: report.fixes.length, warnings: report.warnings.length }, "Preflight completado");

    return report;
  }

  /**
   * Ejecuta un paso y genera checkpoint
   */
  private async step(name: string, fn: () => Promise<void>): Promise<void> {
    this.log.info({ step: name }, `>>> Paso ${name}`);
    await fn();
    await this.hooks.snapshot(name);
    this.log.debug({ step: name }, `Checkpoint ${name} guardado`);
  }
}

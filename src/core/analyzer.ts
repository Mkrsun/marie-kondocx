import { Docx } from "../io/Docx.js";
import type { DocxAnalysis } from "./types.js";

/**
 * Analiza la estructura de un documento DOCX y genera estadísticas detalladas
 */
export async function analyzeDocx(docx: Docx): Promise<DocxAnalysis> {
  const grab = async (p: string) => (await docx.read(p)) ?? "";

  const doc = await grab("word/document.xml");
  const styles = await grab("word/styles.xml");
  const numbering = await grab("word/numbering.xml");
  const comments = await grab("word/comments.xml");
  const relsDoc = await grab("word/_rels/document.xml.rels");
  const contentTypes = await grab("[Content_Types].xml");
  const custom1 = await grab("customXml/item1.xml");

  // Headers y footers para análisis completo
  const headers = docx.list("word/header");
  const footers = docx.list("word/footer");
  const allParts = [doc, ...await Promise.all(headers.map(h => grab(h))), ...await Promise.all(footers.map(f => grab(f)))].join("");

  // Helper para contar ocurrencias de regex
  const count = (xml: string, re: RegExp) => (xml.match(re) ?? []).length;

  // Análisis de custom XML bindings
  const hasBindings = /w:dataBinding\b/i.test(allParts) || /w:storeItemID\b/i.test(allParts);
  const customXmlItems = docx.list("customXml/").filter(p => /item\d+\.xml$/.test(p)).length;

  const stats: DocxAnalysis = {
    parts: {
      document: !!doc,
      styles: !!styles,
      numbering: !!numbering,
      comments: !!comments,
      rels: !!relsDoc,
      customXml: !!custom1,
    },
    comments: {
      entries: count(comments, /<w:comment\b/gi),
      markersStart: count(allParts, /<w:commentRangeStart\b/gi),
      markersEnd: count(allParts, /<w:commentRangeEnd\b/gi),
      refs: count(allParts, /<w:commentReference\b/gi),
    },
    styles: {
      total: count(styles, /<w:style\b/gi),
      defaults: count(styles, /w:default="1"/gi),
      normal: count(styles, /w:styleId="Normal"/g),
      tableNormal: count(styles, /w:styleId="TableNormal"/g),
    },
    numbering: {
      abstractNum: count(numbering, /<w:abstractNum\b/gi),
      num: count(numbering, /<w:num\b/gi),
      lvl: count(numbering, /<w:lvl\b/gi),
    },
    contentTypes: {
      hasWdp: /Extension="wdp"/i.test(contentTypes),
      hasEmf: /Extension="emf"/i.test(contentTypes),
    },
    rels: {
      total: count(relsDoc, /<Relationship\b/gi),
      images: count(relsDoc, /Type="[^"]*image"/gi),
      missingMediaGuess: 0, // Se llenará en rels transform
    },
    altContent: {
      count: count(doc, /<mc:AlternateContent\b/gi),
    },
    revisions: {
      insertions: count(allParts, /<w:ins\b/gi),
      deletions: count(allParts, /<w:del\b/gi),
    },
    sdt: {
      count: count(allParts, /<w:sdt\b/gi),
      withBinding: count(allParts, /<w:dataBinding\b/gi),
    },
    customXml: {
      items: customXmlItems,
      hasBindings,
    }
  };

  return stats;
}

/**
 * Genera un reporte legible en texto del análisis
 */
export function formatAnalysis(analysis: DocxAnalysis, label = "Análisis"): string {
  const lines: string[] = [`\n=== ${label} ===\n`];

  lines.push(`📦 Partes del documento:`);
  lines.push(`   - document.xml: ${analysis.parts.document ? "✓" : "✗"}`);
  lines.push(`   - styles.xml: ${analysis.parts.styles ? "✓" : "✗"}`);
  lines.push(`   - numbering.xml: ${analysis.parts.numbering ? "✓" : "✗"}`);
  lines.push(`   - comments.xml: ${analysis.parts.comments ? "✓" : "✗"}`);
  lines.push(`   - customXml/: ${analysis.parts.customXml ? "✓" : "✗"}`);

  lines.push(`\n💬 Comentarios:`);
  lines.push(`   - Entradas: ${analysis.comments.entries}`);
  lines.push(`   - Marcadores Start: ${analysis.comments.markersStart}`);
  lines.push(`   - Marcadores End: ${analysis.comments.markersEnd}`);
  lines.push(`   - Referencias: ${analysis.comments.refs}`);

  lines.push(`\n🎨 Estilos:`);
  lines.push(`   - Total: ${analysis.styles.total}`);
  lines.push(`   - Defaults: ${analysis.styles.defaults}`);
  lines.push(`   - Normal: ${analysis.styles.normal}`);
  lines.push(`   - TableNormal: ${analysis.styles.tableNormal}`);

  lines.push(`\n🔢 Numeración:`);
  lines.push(`   - AbstractNum: ${analysis.numbering.abstractNum}`);
  lines.push(`   - Num: ${analysis.numbering.num}`);
  lines.push(`   - Niveles: ${analysis.numbering.lvl}`);

  lines.push(`\n🔄 Revisiones (Track Changes):`);
  lines.push(`   - Inserciones: ${analysis.revisions.insertions}`);
  lines.push(`   - Eliminaciones: ${analysis.revisions.deletions}`);

  lines.push(`\n📋 Content Controls (SDT):`);
  lines.push(`   - Total SDT: ${analysis.sdt.count}`);
  lines.push(`   - Con bindings: ${analysis.sdt.withBinding}`);

  lines.push(`\n⚡ Alternate Content:`);
  lines.push(`   - mc:AlternateContent: ${analysis.altContent.count}`);

  lines.push(`\n🔗 Custom XML:`);
  lines.push(`   - Items: ${analysis.customXml.items}`);
  lines.push(`   - Tiene bindings: ${analysis.customXml.hasBindings ? "Sí" : "No"}`);

  lines.push(`\n📎 Content Types:`);
  lines.push(`   - Soporte .wdp: ${analysis.contentTypes.hasWdp ? "✓" : "✗"}`);
  lines.push(`   - Soporte .emf: ${analysis.contentTypes.hasEmf ? "✓" : "✗"}`);

  lines.push(`\n🔗 Relaciones:`);
  lines.push(`   - Total: ${analysis.rels.total}`);
  lines.push(`   - Imágenes: ${analysis.rels.images}`);

  return lines.join("\n");
}

import type { Logger } from "../util/log.js";
import type { PreflightReport } from "../core/types.js";

/**
 * Verifica si un XML de styles necesita saneamiento sin procesarlo completamente
 */
export function needsSanitization(xmlString: string): boolean {
  // Casos que requieren procesamiento:
  // 1. No tiene docDefaults
  if (!xmlString.includes('<w:docDefaults>')) return true;
  
  // 2. Múltiples docDefaults
  const docDefaultsMatches = xmlString.match(/<w:docDefaults>/g);
  if (docDefaultsMatches && docDefaultsMatches.length > 1) return true;
  
  // 3. styleIds con caracteres no ASCII-safe
  if (xmlString.includes('styleId=') && /styleId="[^"]*[^A-Za-z0-9_\-"]/g.test(xmlString)) return true;
  
  // Si no detectamos problemas obvios, no necesita saneamiento
  return false;
}

/**
 * Sanea el archivo de estilos:
 * - Asegura un solo w:docDefaults
 * - Elimina styleId duplicados
 * - Normaliza defaults (solo 1 por tipo)
 * - Limpia IDs a ASCII safe
 * - Filtra hijos inválidos por tipo de estilo
 * - Ordena hijos según esquema
 */
export function sanitize(root: any, log: Logger, report: PreflightReport): any {
  const stylesNode = root.find((n: any) => n["w:styles"]);
  if (!stylesNode) {
    log.warn("No se encontró w:styles");
    return root;
  }

  const styles = stylesNode["w:styles"];
  const children = (styles["#text"] ?? []) as any[];

  // 1) Garantizar un solo docDefaults
  const ddIndices = children
    .map((n, i) => ({ i, has: !!n["w:docDefaults"] }))
    .filter(x => x.has)
    .map(x => x.i);

  if (ddIndices.length > 1) {
    // Eliminar duplicados (conservar el primero)
    ddIndices.slice(1).reverse().forEach(i => children.splice(i, 1));
    report.fixes.push(`Styles: docDefaults duplicados (${ddIndices.length}→1)`);
    log.info({ removed: ddIndices.length - 1 }, "docDefaults duplicados eliminados");
  }

  if (ddIndices.length === 0) {
    // Agregar docDefaults básico
    children.unshift({
      "w:docDefaults": {
        "#text": [
          { "w:rPrDefault": { "#text": [{ "w:rPr": {} }] } },
          { "w:pPrDefault": { "#text": [{ "w:pPr": {} }] } },
        ]
      }
    });
    report.fixes.push("Styles: docDefaults agregado");
    log.info("docDefaults creado");
  }

  // 2) Deduplicar y normalizar estilos
  const seen = new Map<string, number>();
  let deduped = 0;
  let defaultsFixed = 0;

  const sanitized = children.filter(node => {
    if (!node["w:style"]) return true; // No es un estilo, conservar

    const el = node["w:style"];
    const attrs = el;

    // Normalizar styleId a ASCII-safe
    if (attrs["@_w:styleId"] && /[^A-Za-z0-9_-]/.test(attrs["@_w:styleId"])) {
      const old = attrs["@_w:styleId"];
      attrs["@_w:styleId"] = old.replace(/[^A-Za-z0-9_-]/g, "_");
      report.fixes.push(`Styles: styleId normalizado ${old}→${attrs["@_w:styleId"]}`);
      log.debug({ old, new: attrs["@_w:styleId"] }, "styleId normalizado");
    }

    // Solo 1 default por tipo (paragraph, character, table, numbering)
    if (attrs["@_w:default"] === "1") {
      const key = `def:${attrs["@_w:type"]}`;
      if (seen.has(key)) {
        delete attrs["@_w:default"];
        defaultsFixed++;
        log.debug({ type: attrs["@_w:type"] }, "Default duplicado removido");
      } else {
        seen.set(key, 1);
      }
    }

    // Deduplicar por tipo:styleId
    const id = attrs["@_w:styleId"] ?? "";
    const type = attrs["@_w:type"] ?? "paragraph";
    const key = `${type}:${id}`;

    if (seen.has(key)) {
      deduped++;
      log.debug({ styleId: id, type }, "Estilo duplicado eliminado");
      return false;
    }
    seen.set(key, 1);

    // Limpiar hijos inválidos según tipo
    const validChildren = (el["#text"] ?? []).filter((c: any) => {
      const name = Object.keys(c)[0];

      // Character styles no pueden tener w:pPr
      if (type === "character" && name === "w:pPr") {
        log.debug({ styleId: id }, "w:pPr eliminado de character style");
        return false;
      }

      // Table styles no pueden tener w:pPr o w:rPr directamente
      if (type === "table" && (name === "w:pPr" || name === "w:rPr")) {
        log.debug({ styleId: id, child: name }, `${name} eliminado de table style`);
        return false;
      }

      // Numbering styles no pueden tener propiedades de run/table/cell
      if (type === "numbering" && ["w:rPr", "w:tblPr", "w:tcPr", "w:trPr"].includes(name)) {
        log.debug({ styleId: id, child: name }, `${name} eliminado de numbering style`);
        return false;
      }

      return true;
    });

    // Ordenar hijos según esquema
    el["#text"] = sortStyleChildren(validChildren);

    return true;
  });

  styles["#text"] = sanitized;
  report.styles.total = sanitized.filter((n: any) => n["w:style"]).length;
  report.styles.deduped = deduped;
  report.styles.defaultsFixed = defaultsFixed;

  log.info({
    total: report.styles.total,
    deduped,
    defaultsFixed
  }, "Estilos saneados");

  return root;
}

/**
 * Ordena hijos de un estilo según el orden del esquema OOXML
 */
function sortStyleChildren(nodes: any[]): any[] {
  const order = new Map([
    ["w:name", 1],
    ["w:aliases", 2],
    ["w:basedOn", 3],
    ["w:next", 4],
    ["w:link", 5],
    ["w:autoRedefine", 6],
    ["w:hidden", 7],
    ["w:uiPriority", 8],
    ["w:semiHidden", 9],
    ["w:unhideWhenUsed", 10],
    ["w:qFormat", 11],
    ["w:locked", 12],
    ["w:personal", 13],
    ["w:personalCompose", 14],
    ["w:personalReply", 15],
    ["w:rsid", 16],
    ["w:pPr", 20],
    ["w:rPr", 21],
    ["w:tblPr", 22],
    ["w:trPr", 23],
    ["w:tcPr", 24],
    ["w:tblStylePr", 25],
  ]);

  return [...nodes].sort((a, b) => {
    const ka = Object.keys(a)[0];
    const kb = Object.keys(b)[0];
    return (order.get(ka) ?? 50) - (order.get(kb) ?? 50);
  });
}

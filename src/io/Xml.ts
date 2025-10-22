import { XMLBuilder, XMLParser } from "fast-xml-parser";

/**
 * Parser con preserveOrder para roundtrip seguro de namespaces
 */
export const parser = new XMLParser({
  ignoreAttributes: false,
  preserveOrder: true, // Clave para no perder namespaces ni orden
  attributeNamePrefix: "@_",
  parseAttributeValue: false,
  trimValues: false,
  processEntities: false
});

/**
 * Builder que mantiene el orden y namespaces
 * IMPORTANTE: format false para no corromper el XML de DOCX
 */
export const builder = new XMLBuilder({
  ignoreAttributes: false,
  preserveOrder: true,
  attributeNamePrefix: "@_",
  suppressEmptyNode: false,
  format: false,  // No formatear, mantener XML original
  suppressBooleanAttributes: false,
  suppressUnpairedNode: false,
  processEntities: false,
  oneListGroup: false
});

/**
 * Parse XML string a objeto preservando orden
 */
export function parse(xml: string): any {
  return parser.parse(xml);
}

/**
 * Build objeto a XML string
 * Asegura que no se agregue declaración XML automáticamente
 */
export function build(obj: any): string {
  let result = builder.build(obj);
  
  // Remover declaración XML malformada si existe
  if (result.startsWith('<?xml?>')) {
    result = result.substring(7);
  }
  
  // También remover otras declaraciones XML automáticas que podrían aparecer
  if (result.startsWith('<?xml ')) {
    const endIndex = result.indexOf('?>');
    if (endIndex !== -1) {
      result = result.substring(endIndex + 2);
    }
  }
  
  return result;
}

/**
 * Asegura que el namespace w: esté presente en el root
 */
export function ensureWNamespace(root: any): void {
  // root es array como [{ 'w:styles': { '@_xmlns:w': '...', '#text': [...] } }]
  const node = root.find((n: any) => n["w:styles"] || n["w:document"] || n["w:numbering"]);
  if (!node) return;

  const key = Object.keys(node)[0];
  const el = node[key];

  if (!el["@_xmlns:w"]) {
    el["@_xmlns:w"] = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
  }
}

/**
 * Helper para obtener el elemento raíz de un documento parseado
 */
export function getRootElement(parsed: any, tagName: string): any {
  const node = parsed.find((n: any) => n[tagName]);
  return node ? node[tagName] : null;
}

/**
 * Helper para obtener el array de children de un elemento
 */
export function getChildren(element: any): any[] {
  return element?.["#text"] ?? [];
}

/**
 * Helper para setear children de un elemento
 */
export function setChildren(element: any, children: any[]): void {
  element["#text"] = children;
}

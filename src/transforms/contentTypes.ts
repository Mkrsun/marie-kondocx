import type { Logger } from "../util/log.js";
import { Docx } from "../io/Docx.js";

/**
 * Verifica si [Content_Types].xml necesita tipos adicionales
 */
export async function needsContentTypes(docx: Docx): Promise<boolean> {
  const xml = await docx.read("[Content_Types].xml");
  if (!xml) return true; // Si no existe, necesita crearse
  
  // Verificar si faltan tipos de imagen comunes
  const commonTypes = ["png", "jpeg", "gif", "emf", "wmf", "wdp", "svg"];
  for (const ext of commonTypes) {
    if (!new RegExp(`Extension="${ext}"`, "i").test(xml)) {
      return true;
    }
  }
  
  return false;
}

/**
 * Asegura que [Content_Types].xml tenga todos los tipos necesarios
 */
export async function ensure(docx: Docx, log: Logger): Promise<string[]> {
  const xml = (await docx.read("[Content_Types].xml")) ?? '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"></Types>';

  // Defaults que deben existir para formatos de imagen comunes
  const DEFAULTS: Record<string, string> = {
    "png": "image/png",
    "jpg": "image/jpeg",
    "jpeg": "image/jpeg",
    "gif": "image/gif",
    "bmp": "image/bmp",
    "tif": "image/tiff",
    "tiff": "image/tiff",
    "emf": "image/x-emf",
    "wmf": "image/x-wmf",
    "wdp": "image/vnd.ms-photo", // HD Photo
    "svg": "image/svg+xml"
  };

  let changed = xml;
  const added: string[] = [];

  // Asegurar defaults
  for (const [ext, ct] of Object.entries(DEFAULTS)) {
    const already = new RegExp(`<Default[^>]+Extension="${ext}"`, "i").test(xml);
    if (!already) {
      const entry = `<Default Extension="${ext}" ContentType="${ct}"/>`;
      // Insertar antes de </Types>
      changed = changed.replace("</Types>", `  ${entry}\n</Types>`);
      added.push(ext);
    }
  }

  // Asegurar overrides comunes si faltan
  const ensureOverride = (part: string, ct: string) => {
    const re = new RegExp(`<Override[^>]+PartName="/${part}"`, "i");
    if (!re.test(changed)) {
      const entry = `<Override PartName="/${part}" ContentType="${ct}"/>`;
      changed = changed.replace("</Types>", `  ${entry}\n</Types>`);
      log.debug({ part }, "Override agregado");
    }
  };

  // Estos son opcionales, solo agregamos si vemos que hacen falta
  // (el template debería tenerlos, pero por si acaso)

  if (changed !== xml) {
    docx.write("[Content_Types].xml", changed);
    log.info({ added }, "Content types actualizados");
  }

  return added;
}

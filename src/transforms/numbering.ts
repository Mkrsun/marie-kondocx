import type { Logger } from "../util/log.js";
import type { PreflightReport } from "../core/types.js";

/**
 * Verifica si un XML de numbering necesita limpieza sin procesarlo completamente
 */
export function needsCleaning(xmlString: string): boolean {
  // Si no hay contenido de numeración, no necesita limpieza
  if (!xmlString.includes('<w:abstractNum') && !xmlString.includes('<w:num>')) {
    return false;
  }
  
  // Para simplificar, si hay contenido de numeración, asumimos que puede necesitar limpieza
  // Una implementación más sofisticada podría hacer análisis más detallado
  return xmlString.includes('<w:abstractNum') || xmlString.includes('<w:num>');
}

/**
 * Limpia y normaliza el archivo de numeración:
 * - Elimina w:num y w:abstractNum huérfanos (sin referencias)
 * - Normaliza niveles (w:lvl ilvl debe ser 0-8, sin duplicados)
 * - Valida estructura de listas
 */
export function cleanAndNormalize(root: any, log: Logger, report: PreflightReport): any {
  const numNode = root.find((n: any) => n["w:numbering"]);
  if (!numNode) {
    log.debug("No se encontró w:numbering");
    return root;
  }

  const numbering = numNode["w:numbering"];
  const kids = (numbering["#text"] ?? []) as any[];

  // Separar abstractNum y num
  const abstractNums = kids.filter((n: any) => n["w:abstractNum"]).map((n: any) => n["w:abstractNum"]);
  const nums = kids.filter((n: any) => n["w:num"]).map((n: any) => n["w:num"]);

  // Construir sets de IDs
  const abstractIds = new Set(
    abstractNums.map((a: any) => String(a["@_w:abstractNumId"] ?? ""))
  );
  const usedAbstractIds = new Set<string>();
  const keepKids: any[] = [];

  log.debug({
    abstractNums: abstractNums.length,
    nums: nums.length
  }, "Analizando numeración");

  // 1) Procesar w:num y detectar cuáles abstractNum se usan
  for (const n of nums) {
    const children = (n["#text"] ?? []) as any[];
    const aRefNode = children.find((c: any) => c["w:abstractNumId"]);
    const id = String(aRefNode?.["w:abstractNumId"]?.["@_w:val"] ?? "");

    if (abstractIds.has(id)) {
      keepKids.push({ "w:num": n });
      usedAbstractIds.add(id);
    } else {
      report.numbering.fixes++;
      log.warn({ numId: n["@_w:numId"], abstractRef: id }, "w:num huérfano eliminado (referencia inválida)");
    }
  }

  // 2) Procesar abstractNum: conservar solo los usados (o todos si no hay nums)
  const hasNums = keepKids.some(n => n["w:num"]);

  for (const a of abstractNums) {
    const id = String(a["@_w:abstractNumId"] ?? "");

    // Si no hay nums, conservar todos los abstractNum
    // Si hay nums, conservar solo los referenciados
    if (!hasNums || usedAbstractIds.has(id)) {
      // Normalizar niveles
      const normalized = normalizeLevels(a, log, report);
      keepKids.unshift({ "w:abstractNum": normalized });
    } else {
      report.numbering.fixes++;
      log.warn({ abstractNumId: id }, "w:abstractNum sin uso eliminado");
    }
  }

  // Actualizar hijos
  numbering["#text"] = keepKids;

  log.info({
    kept: keepKids.length,
    fixes: report.numbering.fixes
  }, "Numeración limpiada");

  return root;
}

/**
 * Normaliza los niveles de un abstractNum:
 * - ilvl debe estar entre 0-8
 * - No debe haber niveles duplicados
 * - Ordena niveles por ilvl
 */
function normalizeLevels(abstractNum: any, log: Logger, report: PreflightReport): any {
  const children = (abstractNum["#text"] ?? []) as any[];
  const levels = children.filter((c: any) => c["w:lvl"]).map((c: any) => c["w:lvl"]);
  const nonLevels = children.filter((c: any) => !c["w:lvl"]);

  const seen = new Set<number>();
  const validLevels: any[] = [];

  for (const lvl of levels) {
    const ilvl = Number(lvl["@_w:ilvl"] ?? "0");

    // Validar rango
    if (ilvl < 0 || ilvl > 8) {
      report.numbering.fixes++;
      log.warn({
        abstractNumId: abstractNum["@_w:abstractNumId"],
        ilvl
      }, "Nivel fuera de rango (0-8) eliminado");
      continue;
    }

    // Validar duplicados
    if (seen.has(ilvl)) {
      report.numbering.fixes++;
      log.warn({
        abstractNumId: abstractNum["@_w:abstractNumId"],
        ilvl
      }, "Nivel duplicado eliminado");
      continue;
    }

    seen.add(ilvl);
    validLevels.push(lvl);
  }

  // Ordenar por ilvl
  validLevels.sort((a, b) => {
    const aLvl = Number(a["@_w:ilvl"] ?? "0");
    const bLvl = Number(b["@_w:ilvl"] ?? "0");
    return aLvl - bLvl;
  });

  // Reconstruir hijos: primero los no-level, luego los levels ordenados
  abstractNum["#text"] = [
    ...nonLevels,
    ...validLevels.map(lvl => ({ "w:lvl": lvl }))
  ];

  return abstractNum;
}

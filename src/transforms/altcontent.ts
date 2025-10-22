import type { Logger } from "../util/log.js";
import type { PreflightReport } from "../core/types.js";

/**
 * Resuelve mc:AlternateContent reemplazándolo por su mc:Fallback
 * Esto elimina contenido específico de versiones nuevas de Word y deja solo el fallback compatible
 * 
 * PROCESAMIENTO DIRECTO DE XML COMO STRING - NO USA fast-xml-parser
 */
export function resolveFallbacks(xmlString: string, log: Logger, report: PreflightReport): string {
  let count = 0;
  let processedCount = 0;
  let choiceCount = 0;
  let fallbackCount = 0;

  log.info("🔄 Procesando mc:AlternateContent directamente en XML string");

  // Regex para encontrar elementos mc:AlternateContent completos
  const alternateContentRegex = /<mc:AlternateContent[^>]*>[\s\S]*?<\/mc:AlternateContent>/g;
  
  let result = xmlString;
  let match;
  
  // Resetear el regex para usar exec() correctamente
  alternateContentRegex.lastIndex = 0;
  
  while ((match = alternateContentRegex.exec(xmlString)) !== null) {
    count++;
    const fullElement = match[0];
    
    log.info(`🔍 Procesando mc:AlternateContent #${count}: ${fullElement.substring(0, 100)}...`);
    
    // Buscar mc:Choice primero (contenido moderno preferido)
    const choiceMatch = fullElement.match(/<mc:Choice[^>]*>([\s\S]*?)<\/mc:Choice>/);
    const fallbackMatch = fullElement.match(/<mc:Fallback[^>]*>([\s\S]*?)<\/mc:Fallback>/);
    
    let replacement = fullElement; // Por defecto conservar original
    
    if (choiceMatch) {
      choiceCount++;
      replacement = choiceMatch[1];
      log.info(`🎯 Choice encontrado - usando Choice content: ${replacement.substring(0, 50)}...`);
      processedCount++;
    } else if (fallbackMatch) {
      fallbackCount++;
      replacement = fallbackMatch[1];
      log.info(`🔄 Fallback encontrado - usando Fallback content: ${replacement.substring(0, 50)}...`);
      processedCount++;
    } else {
      log.warn("⚠️ mc:AlternateContent sin Choice/Fallback válido - CONSERVANDO ORIGINAL");
    }
    
    // Reemplazar en el resultado final
    result = result.replace(fullElement, replacement);
  }

  if (count > 0) {
    report.fixes.push(`AlternateContent: ${processedCount}/${count} procesados (Choice: ${choiceCount}, Fallback: ${fallbackCount})`);
    log.info({ 
      total: count, 
      processed: processedCount, 
      choice: choiceCount, 
      fallback: fallbackCount 
    }, "🎯 AlternateContent procesado con detalles");
  } else {
    log.info("No se encontraron elementos mc:AlternateContent para procesar");
  }

  return result;
}

/**
 * Verifica si un XML contiene mc:AlternateContent sin procesarlo completamente
 */
export function hasAlternateContent(xmlContent: string): boolean {
    // Check for actual XML elements, not just text content containing "mc:"
    return /<mc:AlternateContent\b/.test(xmlContent);
}

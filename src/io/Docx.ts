import JSZip from "jszip";
import { readFile, writeFile } from "node:fs/promises";

/**
 * Wrapper sobre JSZip para manejar archivos .docx
 */
export class Docx {
  private zip: JSZip;

  private constructor(zip: JSZip) {
    this.zip = zip;
  }

  /**
   * Abre un archivo .docx desde disco
   */
  static async open(path: string): Promise<Docx> {
    const buff = await readFile(path);
    const zip = await JSZip.loadAsync(buff);
    return new Docx(zip);
  }

  /**
   * Lee un archivo interno del .docx
   */
  async read(path: string): Promise<string | null> {
    const f = this.zip.file(path);
    return f ? await f.async("string") : null;
  }

  /**
   * Escribe un archivo interno del .docx
   */
  write(path: string, content: string): void {
    this.zip.file(path, content);
  }

  /**
   * Elimina un archivo interno del .docx
   */
  delete(path: string): void {
    this.zip.remove(path);
  }

  /**
   * Verifica si existe un archivo interno
   */
  exists(path: string): boolean {
    return !!this.zip.file(path);
  }

  /**
   * Lista archivos internos con un prefijo dado
   */
  list(prefix = "word/"): string[] {
    return Object.keys(this.zip.files).filter(p => p.startsWith(prefix) && !p.endsWith("/"));
  }

  /**
   * Guarda el .docx a disco
   */
  async saveAs(path: string): Promise<void> {
    const buff = await this.zip.generateAsync({
      type: "nodebuffer",
      compression: "DEFLATE",
      compressionOptions: { level: 9 }
    });
    await writeFile(path, buff);
  }
}

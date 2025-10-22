# marie-kondocx 🧹

**DOCX Preflight & Clean** - Análisis y limpieza incremental de documentos Word para templating y merge

> "Does this XML spark joy?" - Marie Kondo, probablemente

## 🎯 ¿Qué hace?

**marie-kondocx** prepara documentos `.docx` para uso como templates, limpiándolos de todo el "ruido" que Word acumula:

- ✅ **AlternateContent** → Resuelve a Fallback (compatibilidad)
- ✅ **Track Changes** → Acepta revisiones (w:ins/w:del)
- ✅ **Content Controls (SDT)** → Aplana controles para texto plano
- ✅ **Estilos** → Deduplica, normaliza defaults, limpia IDs
- ✅ **Numeración** → Elimina huérfanos, normaliza niveles
- ✅ **Comentarios** → Elimina completamente (opcional)
- ✅ **Relaciones** → Valida y limpia referencias rotas
- ✅ **Content Types** → Asegura soporte para formatos de imagen (incluye .wdp)
- ✅ **CustomXML** → Política configurable (keep/remove/auto)
- ✅ **Secciones** → Asegura sectPr válido
- ✅ **Texto** → Desfragmenta runs para preservar placeholders como `{{nombre}}`

## 🚀 Características únicas

### 📸 Checkpoints por paso

Cada transformación genera:
- **Un DOCX** con el estado después del paso (`steps/input.01-altcontent.docx`)
- **Un JSON** con análisis detallado (`analysis/input.after-step-01-altcontent.json`)

Esto permite:
- Ver exactamente **qué cambió** en cada paso
- **Debugging preciso** si algo falla
- **Auditar** la limpieza de cada template

### 📊 Análisis antes/después

Genera reportes JSON detallados con:
- Conteo de comentarios, estilos, numeraciones
- Track changes detectados
- SDT con/sin bindings
- CustomXML y sus usos
- Relaciones y media

### 🎛️ Políticas configurables

- `--keepComments` : Conservar comentarios
- `--flattenSDT` : Aplanar content controls
- `--acceptRevisions` : Aceptar track changes
- `--customXml auto|keep|remove` : Política de CustomXML

## 📦 Instalación

```bash
npm install
npm run build
```

## 🎬 Uso

### Comando básico

```bash
node dist/cli.js run -i input.docx -o output.docx
```

### Con todas las opciones

```bash
node dist/cli.js run \
  -i ./templates/contrato.docx \
  -o ./clean/contrato-clean.docx \
  --report ./reports/contrato-report.json \
  --stepsDir ./steps \
  --analysisDir ./analysis \
  --customXml auto \
  --acceptRevisions true \
  --flattenSDT true \
  --verbose
```

### Flags disponibles

| Flag | Tipo | Default | Descripción |
|------|------|---------|-------------|
| `-i, --input` | string | **requerido** | Archivo .docx de entrada |
| `-o, --output` | string | **requerido** | Archivo .docx limpio de salida |
| `-r, --report` | string | `report.json` | Archivo de reporte JSON |
| `--stepsDir` | string | `steps` | Directorio para DOCX por paso |
| `--analysisDir` | string | `analysis` | Directorio para análisis JSON |
| `--keepComments` | boolean | `false` | No eliminar comentarios |
| `--flattenSDT` | boolean | `true` | Aplanar content controls |
| `--acceptRevisions` | boolean | `true` | Aceptar track changes |
| `--customXml` | `keep\|remove\|auto` | `auto` | Política de CustomXML |
| `-v, --verbose` | boolean | `false` | Logging detallado |

## 📁 Estructura de salida

Después de ejecutar, tendrás:

```
steps/
  input.00-original.docx          ← Copia del original
  input.01-altcontent.docx        ← Después de resolver AlternateContent
  input.02-revisions.docx         ← Después de aceptar track changes
  input.03-sdt.docx               ← Después de aplanar SDT
  input.04-text.docx              ← Después de desfragmentar texto
  input.05-styles.docx            ← Después de sanear estilos
  input.06-numbering.docx         ← Después de limpiar numeración
  input.07-comments.docx          ← Después de eliminar comentarios
  input.08-rels.docx              ← Después de validar relaciones
  input.09-contenttypes.docx      ← Después de actualizar content types
  input.10-sections.docx          ← Después de asegurar sectPr
  input.11-customxml.docx         ← Después de aplicar política customXML
  input.99-final.docx             ← Final (igual a output)

analysis/
  input.before.json               ← Análisis inicial
  input.after-step-01-altcontent.json
  input.after-step-02-revisions.json
  ... (un JSON por paso)
  input.after.json                ← Análisis final

report.json                       ← Reporte de fixes y warnings
output.docx                       ← Documento limpio final
```

## 🔍 CustomXML: ¿Qué es y cuándo importa?

### ¿Qué es CustomXML?

Carpeta `/customXml/` con archivos XML que guardan **datos** para enlazar con **controles de contenido (SDT)** usando bindings.

### ¿Cuándo es necesario?

✅ **Mantenerlo** si:
- Tu template usa **bindings** (`w:dataBinding`, `w:storeItemID`)
- Tienes SDT que se "alimentan" de datos XML
- Ejemplo: formularios con campos conectados a fuente de datos

❌ **Eliminarlo** si:
- Usas placeholders de texto (`{{nombre}}`, `${rut}`)
- Ya aplanaste todos los SDT
- No hay bindings activos

### Política `auto`

Detecta automáticamente si hay bindings:
- **Con bindings** → mantiene CustomXML
- **Sin bindings** → elimina CustomXML

## 📊 Ejemplo de análisis

```json
{
  "parts": {
    "document": true,
    "styles": true,
    "comments": true,
    "customXml": false
  },
  "comments": {
    "entries": 10,
    "markersStart": 10,
    "markersEnd": 10,
    "refs": 10
  },
  "styles": {
    "total": 45,
    "defaults": 2,
    "normal": 1
  },
  "revisions": {
    "insertions": 5,
    "deletions": 3
  },
  "sdt": {
    "count": 8,
    "withBinding": 0
  }
}
```

## 🧪 Ejemplo de reporte

```json
{
  "fixes": [
    "AlternateContent: 3 elementos resueltos",
    "Revisions: 5 inserciones aceptadas, 3 eliminaciones aceptadas",
    "SDT: 8 controles aplanados",
    "Text: 12 runs desfragmentados",
    "Styles: docDefaults duplicados (2→1)",
    "Numbering: w:num huérfano eliminado"
  ],
  "warnings": [],
  "styles": {
    "total": 43,
    "deduped": 2,
    "defaultsFixed": 1
  },
  "comments": {
    "removed": 10,
    "kept": 0
  },
  "numbering": {
    "fixes": 3
  },
  "contentTypes": {
    "added": ["wdp", "svg"]
  },
  "customXml": {
    "action": "removed",
    "hasBindings": false
  }
}
```

## 🏗️ Arquitectura

### Principios SOLID (ligero)

- **S (Single Responsibility)**: Cada transform hace una cosa
- **O (Open/Closed)**: Agregar transforms sin tocar existentes
- **D (Dependency Inversion)**: Logging y I/O inyectados

### Módulos principales

```
src/
  cli.ts                    ← CLI con yargs
  index.ts                  ← Orquestador principal
  core/
    PreflightService.ts     ← Pipeline con checkpoints
    analyzer.ts             ← Análisis de estructura DOCX
    types.ts                ← Tipos TypeScript
  io/
    Docx.ts                 ← Wrapper sobre JSZip
    Xml.ts                  ← Parser/Builder con preserveOrder
  transforms/               ← Un archivo por transformación
    altcontent.ts
    revisions.ts
    sdt.ts
    text.ts
    styles.ts
    numbering.ts           ← ¡Limpia huérfanos!
    comments.ts
    rels.ts
    contentTypes.ts
    sections.ts
    customxml.ts
  util/
    log.ts                  ← Logger con pino
```

## 🛠️ Stack tecnológico

- **JSZip** (MIT) - ZIP/Unzip de DOCX
- **fast-xml-parser** (MIT) - Parse/Build XML con `preserveOrder`
- **pino** (MIT) - Logging estructurado
- **yargs** (MIT) - CLI parsing
- **TypeScript** - Type safety

## 🎯 Casos de uso

### 1. Templates para merge

```bash
# Limpiar template antes de usar para mail merge
node dist/cli.js run -i template-contrato.docx -o clean-template.docx
```

### 2. Múltiples templates con batch

```bash
for file in templates/*.docx; do
  name=$(basename "$file" .docx)
  node dist/cli.js run -i "$file" -o "clean/$name-clean.docx" --stepsDir "steps/$name" --analysisDir "analysis/$name"
done
```

### 3. Debugging de template corrupto

Si un template da "contenido no legible", revisa:
1. `analysis/template.before.json` → qué tiene
2. `steps/template.*.docx` → en qué paso falla
3. Logs con `-v` → detalles del error

## ⚠️ Limitaciones conocidas

- **Macros (VBA)**: No se procesan, se conservan tal cual
- **Objetos OLE**: Se mantienen sin modificar
- **Formas complejas**: SmartArt, diagramas - se conservan
- **Fuentes embebidas**: No se optimizan

## 🔜 Próximas mejoras

- [ ] Merge CLI (combinar múltiples DOCX limpios)
- [ ] Validador de placeholders (detectar `{{variables}}`)
- [ ] Optimizador de media (comprimir imágenes)
- [ ] Diff visual entre pasos
- [ ] Unit tests con snapshots
- [ ] Docker image

## 📄 Licencia

MIT

## 🤝 Contribuir

PRs bienvenidos! Para cambios grandes, abre primero un issue.

## 🐛 Reportar bugs

[GitHub Issues](https://github.com/tu-usuario/marie-kondocx/issues)

## 💡 ¿Por qué "marie-kondocx"?

Porque limpia tu DOCX dejando solo lo que "spark joy" (y funciona) ✨

---

**Hecho con ❤️ para hacer el trabajo con templates de Word menos doloroso**

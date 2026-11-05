export const NS = {
  w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  wp: 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
  ct: 'http://schemas.openxmlformats.org/package/2006/content-types',
  rr: 'http://schemas.openxmlformats.org/package/2006/relationships',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  xml: 'http://www.w3.org/XML/1998/namespace',
} as const;

export type PartCategory = 'main' | 'header_footer' | 'notes' | 'settings';

/** Content types (from [Content_Types].xml overrides) that the merge engine cares about. */
export const CONTENT_TYPE_CATEGORIES: Readonly<Record<string, PartCategory>> = {
  'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml': 'main',
  'application/vnd.ms-word.document.macroEnabled.main+xml': 'main',
  'application/vnd.openxmlformats-officedocument.wordprocessingml.template.main+xml': 'main',
  'application/vnd.ms-word.template.macroEnabledTemplate.main+xml': 'main',
  'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml': 'header_footer',
  'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml': 'header_footer',
  'application/vnd.openxmlformats-officedocument.wordprocessingml.footnotes+xml': 'notes',
  'application/vnd.openxmlformats-officedocument.wordprocessingml.endnotes+xml': 'notes',
  'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml': 'settings',
};

export const SEPARATORS = [
  'page_break',
  'column_break',
  'textWrapping_break',
  'continuous_section',
  'evenPage_section',
  'nextColumn_section',
  'nextPage_section',
  'oddPage_section',
] as const;

export type Separator = (typeof SEPARATORS)[number];

/** Zip entries that are already compressed: stored as-is instead of being deflated again. */
export const STORED_EXTENSIONS: ReadonlySet<string> = new Set([
  'png',
  'jpg',
  'jpeg',
  'jfif',
  'gif',
  'webp',
  'wdp',
  'emz',
  'wmz',
  'svgz',
  'mp3',
  'mp4',
  'm4a',
  'mov',
  'avi',
  'wmv',
  'zip',
  'docx',
  'xlsx',
  'pptx',
  'odttf',
]);

/** w:sectPr children that must come after w:type (ECMA-376 CT_SectPr sequence). */
export const SECTPR_AFTER_TYPE: ReadonlySet<string> = new Set([
  'pgSz',
  'pgMar',
  'paperSrc',
  'pgBorders',
  'lnNumType',
  'pgNumType',
  'cols',
  'formProt',
  'vAlign',
  'noEndnote',
  'titlePg',
  'textDirection',
  'bidi',
  'rtlGutter',
  'docGrid',
  'printerSettings',
  'sectPrChange',
]);

/** w:settings children that must come after w:updateFields (ECMA-376 CT_Settings sequence). */
export const SETTINGS_AFTER_UPDATE_FIELDS: ReadonlySet<string> = new Set([
  'hdrShapeDefaults',
  'footnotePr',
  'endnotePr',
  'compat',
  'docVars',
  'rsids',
  'mathPr',
  'attachedSchema',
  'themeFontLang',
  'clrSchemeMapping',
  'doNotIncludeSubdocsInStats',
  'doNotAutoCompressPictures',
  'forceUpgrade',
  'captions',
  'readModeInkLockDown',
  'smartTagType',
  'schemaLibrary',
  'shapeDefaults',
  'doNotEmbedSmartTags',
  'decimalSymbol',
  'listSeparator',
]);

/** Tag name of the internal placeholder element that replaces a field in the XML tree. */
export const PLACEHOLDER = 'MergeField';

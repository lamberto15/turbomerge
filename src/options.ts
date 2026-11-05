import type { Separator } from './constants.ts';

export type { Separator };

/** What happens to fields that have no data when the document is written. */
export type KeepFields = 'none' | 'some' | 'all';

/** Whether to ask Word to refresh all fields when the generated document is opened. */
export type AutoUpdateFields = 'no' | 'auto' | 'always';

/** A record of merge data. Arrays of records fill table rows; nested objects are reachable with dot paths. */
export type MergeRecord = Readonly<Record<string, unknown>>;

export interface FormatterContext {
  /** Field name the value belongs to. */
  readonly field: string;
  /** BCP 47 locale in effect. */
  readonly locale: string;
  readonly timeZone: string | undefined;
}

/** Custom `\* name` text formatter. */
export type Formatter = (value: unknown, context: FormatterContext) => unknown;

export interface MailMergeWarning {
  readonly code: 'INVALID_FORMAT' | 'INVALID_INSTRUCTION' | 'INVALID_CONDITION';
  readonly message: string;
  readonly field?: string;
}

export interface MailMergeOptions {
  /** Remove tables whose anchor received an empty array. Default `false`. */
  removeEmptyTables?: boolean;
  /** `'auto'` sets Word's update-fields-on-open flag only when nested fields exist. Default `'no'`. */
  autoUpdateFieldsOnOpen?: AutoUpdateFields;
  /**
   * `'none'` (default): every field is replaced; fields without data become `emptyValue`.
   * `'some'`: fields without data are kept as live Word fields.
   * `'all'`: every field is kept as a live Word field, with its result updated when data exists.
   */
  keepFields?: KeepFields;
  /** Evaluate `IF` fields. Default `false` (left for Word to compute). */
  mergeIfFields?: boolean;
  /** Overwrite existing table rows instead of inserting new ones (throws on overflow). Default `false`. */
  tableRowsReplaceMode?: boolean;
  /** Locale used for dates, month/day names and default number rendering. Default `'en-US'`. */
  locale?: string;
  /** IANA time zone used to render `Date` values. Default: the process time zone. */
  timeZone?: string;
  /** Custom formatters invoked by `\* name` switches (case-insensitive). */
  formatters?: Readonly<Record<string, Formatter>>;
  /** Throw `MissingFieldsError` from `render*()` when a record lacks a merge field. Default `false`. */
  strict?: boolean;
  /** Receives non-fatal problems. Default: `process.emitWarning`. */
  onWarning?: (warning: MailMergeWarning) => void;
}

export interface ResolvedOptions {
  readonly removeEmptyTables: boolean;
  readonly autoUpdateFieldsOnOpen: AutoUpdateFields;
  readonly keepFields: KeepFields;
  readonly mergeIfFields: boolean;
  readonly tableRowsReplaceMode: boolean;
  readonly locale: string;
  readonly timeZone: string | undefined;
  readonly formatters: ReadonlyMap<string, Formatter>;
  readonly strict: boolean;
  readonly onWarning: (warning: MailMergeWarning) => void;
}

const defaultWarning = (warning: MailMergeWarning): void => {
  process.emitWarning(warning.message, { type: 'MailMergeWarning', code: warning.code });
};

export function resolveOptions(options: MailMergeOptions = {}): ResolvedOptions {
  const keepFields = options.keepFields ?? 'none';
  if (!['none', 'some', 'all'].includes(keepFields)) throw new TypeError(`Invalid keepFields: ${keepFields}`);
  const auto = options.autoUpdateFieldsOnOpen ?? 'no';
  if (!['no', 'auto', 'always'].includes(auto)) throw new TypeError(`Invalid autoUpdateFieldsOnOpen: ${auto}`);
  const locale = options.locale ?? 'en-US';
  // Fail fast on bad locale / time zone instead of in the middle of a merge.
  new Intl.DateTimeFormat(locale, { timeZone: options.timeZone });

  return Object.freeze({
    removeEmptyTables: options.removeEmptyTables ?? false,
    autoUpdateFieldsOnOpen: auto,
    keepFields,
    mergeIfFields: options.mergeIfFields ?? false,
    tableRowsReplaceMode: options.tableRowsReplaceMode ?? false,
    locale,
    timeZone: options.timeZone,
    formatters: new Map(Object.entries(options.formatters ?? {}).map(([k, v]) => [k.toLowerCase(), v])),
    strict: options.strict ?? false,
    onWarning: options.onWarning ?? defaultWarning,
  });
}

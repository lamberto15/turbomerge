import type { FormatterContext, ResolvedOptions } from './options.ts';

/**
 * Implements the Word field switches:
 *   \# numeric picture   \@ date-time picture   \* text format
 *   \b text before        \f text after
 * https://support.microsoft.com/office/format-field-results-baa61f5a-5636-4f11-ab4f-6c36ae43508c
 */

const NUMBER_PICTURE = /^([^0.,'#PN]+)?(P\d+|N\d+|[0.,'#]+%?)([^0.,'#%].*)?/s;
const DATE_TOKEN = /'[^']*'|am\/pm|AM\/PM|y+|Y+|M+|d+|D+|h+|H+|m+|s+|S+/g;
const ISO_DATE = /^\d{4}-\d{2}-\d{2}(?:[T ]\d{2}:\d{2}(?::\d{2}(?:\.\d+)?)?(?:Z|[+-]\d{2}:?\d{2})?)?$/;
const NUMERIC = /^[+-]?(?:\d+\.?\d*|\.\d+)(?:e[+-]?\d+)?$/i;

export function* walkSwitches(tokens: readonly string[]): Generator<[flag: string, argument: string]> {
  let i = 0;
  while (i < tokens.length) {
    const token = tokens[i]!;
    const flag = token.slice(0, 2);
    if (!flag) {
      i += 1;
      continue;
    }
    if (token.length > 2) {
      yield [flag, token.slice(2)];
      i += 1;
    } else {
      if (i + 1 >= tokens.length) throw new Error(`Missing argument for switch ${flag}`);
      yield [flag, tokens[i + 1]!];
      i += 2;
    }
  }
}

export interface FormatContext extends FormatterContext {
  readonly options: ResolvedOptions;
}

/** Applies all switches in order; throws on malformed pictures (caller keeps the raw value). */
export function applySwitches(value: unknown, switches: readonly string[], ctx: FormatContext): unknown {
  for (const [flag, arg] of walkSwitches(switches)) {
    switch (flag) {
      case '\\b':
      case '\\f':
        value = isTruthy(value) ? (flag === '\\b' ? arg + toText(value, ctx) : toText(value, ctx) + arg) : null;
        break;
      case '\\#':
        value = formatNumber(value, arg);
        break;
      case '\\@':
        value = formatDate(value, arg, ctx);
        break;
      case '\\*':
        value = formatText(value, arg, ctx);
        break;
    }
  }
  return value;
}

/** Final conversion of a merge value to the text placed in the document. */
export function toText(value: unknown, ctx: FormatterContext): string {
  if (value == null) return '';
  if (value instanceof Date) return defaultDate(value, ctx);
  return String(value);
}

function isTruthy(value: unknown): boolean {
  return !(value == null || value === '' || value === 0 || value === false || value === 0n);
}

export function toNumber(value: unknown): number {
  if (value == null) return 0;
  if (typeof value === 'number') return value;
  if (typeof value === 'bigint' || typeof value === 'boolean') return Number(value);
  if (typeof value === 'string' && NUMERIC.test(value.trim())) return Number(value.trim());
  throw new TypeError(`Value <${String(value)}> is not a number`);
}

export function isNumeric(value: string): boolean {
  return NUMERIC.test(value.trim());
}

function group(intPart: string, sep: string): string {
  return sep ? intPart.replace(/\B(?=(\d{3})+(?!\d))/g, sep) : intPart;
}

export function formatNumber(value: unknown, picture: string): string {
  const match = NUMBER_PICTURE.exec(picture);
  if (!match) throw new Error(`Non conforming number format <${picture}>`);
  const prefix = match[1] ?? '';
  let pattern = match[2]!;
  const suffix = match[3] ?? '';
  const num = toNumber(value);

  if (pattern[0] === 'P') return `${prefix}${(num * 100).toFixed(Number(pattern.slice(1)))}%${suffix}`;
  if (pattern[0] === 'N') return `${prefix}${num.toFixed(Number(pattern.slice(1)))}${suffix}`;
  if (pattern.endsWith('%')) return `${prefix}${(num * 100).toFixed(0)}%${suffix}`;

  const thousands = pattern.includes("'") ? "'" : pattern.includes(',') ? ',' : '';
  pattern = pattern.replace(/[,']/g, '');
  const [digits = '', decimals = ''] = pattern.split('.');
  const minIntDigits = digits.replace(/#/g, '').length;

  const fixed = Math.abs(num).toFixed(decimals.length);
  const [intPart = '0', fraction] = fixed.split('.');
  const paddedInt = minIntDigits > 1 ? intPart.padStart(minIntDigits, '0') : intPart;
  const sign = num < 0 && Number(fixed) !== 0 ? '-' : '';
  return `${prefix}${sign}${group(paddedInt, thousands)}${fraction === undefined ? '' : `.${fraction}`}${suffix}`;
}

// ---- dates -------------------------------------------------------------------------------

const intlCache = new Map<string, Intl.DateTimeFormat>();
function dtf(locale: string, timeZone: string | undefined, opts: Intl.DateTimeFormatOptions): Intl.DateTimeFormat {
  const key = `${locale}|${timeZone ?? ''}|${JSON.stringify(opts)}`;
  let f = intlCache.get(key);
  if (!f) {
    f = new Intl.DateTimeFormat(locale, { ...opts, timeZone });
    intlCache.set(key, f);
  }
  return f;
}

interface DateParts {
  year: number;
  month: number;
  day: number;
  hour: number;
  minute: number;
  second: number;
}

function dateParts(date: Date, timeZone: string | undefined): DateParts {
  const parts = dtf('en-US', timeZone, {
    year: 'numeric', month: 'numeric', day: 'numeric',
    hour: 'numeric', minute: 'numeric', second: 'numeric', hourCycle: 'h23',
  }).formatToParts(date);
  const get = (type: string): number => Number(parts.find((p) => p.type === type)?.value ?? 0);
  return { year: get('year'), month: get('month'), day: get('day'), hour: get('hour'), minute: get('minute'), second: get('second') };
}

/** Accepts Date, epoch-free ISO 8601 strings ("2024-03-09", "2024-03-09T17:07:08Z"). */
export function toDate(value: unknown): Date | null {
  if (value instanceof Date) return Number.isNaN(value.getTime()) ? null : value;
  if (typeof value === 'string' && ISO_DATE.test(value.trim())) {
    const s = value.trim();
    if (s.length === 10) {
      const [y, m, d] = s.split('-').map(Number) as [number, number, number];
      return new Date(y, m - 1, d); // date-only strings are local dates, not UTC midnight
    }
    const date = new Date(s.replace(' ', 'T'));
    return Number.isNaN(date.getTime()) ? null : date;
  }
  return null;
}

const pad = (n: number, len = 2): string => String(n).padStart(len, '0');

export function formatDate(value: unknown, picture: string, ctx: FormatterContext): string {
  if (value == null || value === '') return '';
  const date = toDate(value);
  if (!date) return String(value);
  const p = dateParts(date, ctx.timeZone);
  const name = (opts: Intl.DateTimeFormatOptions): string => dtf(ctx.locale, ctx.timeZone, opts).format(date);

  return picture.replace(DATE_TOKEN, (tok) => {
    if (tok.startsWith("'")) return tok.slice(1, -1);
    if (tok === 'am/pm' || tok === 'AM/PM') return p.hour < 12 ? 'AM' : 'PM';
    const n = tok.length;
    switch (tok[0]) {
      case 'M':
        return n === 1 ? String(p.month) : n === 2 ? pad(p.month) : name({ month: n === 3 ? 'short' : 'long' });
      case 'd':
      case 'D':
        return n === 1 ? String(p.day) : n === 2 ? pad(p.day) : name({ weekday: n === 3 ? 'short' : 'long' });
      case 'y':
      case 'Y':
        return n <= 2 ? pad(p.year % 100) : String(p.year);
      case 'h': {
        const h12 = p.hour % 12 || 12;
        return n === 1 ? String(h12) : pad(h12);
      }
      case 'H':
        return n === 1 ? String(p.hour) : pad(p.hour);
      case 'm':
        return n === 1 ? String(p.minute) : pad(p.minute);
      default: // s / S
        return n === 1 ? String(p.second) : pad(p.second);
    }
  });
}

function defaultDate(date: Date, ctx: FormatterContext): string {
  const p = dateParts(date, ctx.timeZone);
  const dateOnly = p.hour === 0 && p.minute === 0 && p.second === 0;
  return dtf(ctx.locale, ctx.timeZone, dateOnly
    ? { year: 'numeric', month: '2-digit', day: '2-digit' }
    : { year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', second: '2-digit' },
  ).format(date);
}

// ---- text --------------------------------------------------------------------------------

export function formatText(value: unknown, option: string, ctx: FormatContext): unknown {
  const key = option.toLowerCase();
  const custom = ctx.options.formatters.get(key);
  if (custom) return custom(value, ctx);
  const text = (): string => toText(value, ctx);
  switch (key) {
    case 'upper':
      return text().toLocaleUpperCase(ctx.locale);
    case 'lower':
      return text().toLocaleLowerCase(ctx.locale);
    case 'caps':
      return text()
        .toLocaleLowerCase(ctx.locale)
        .replace(/(^|\s)(\p{L})/gu, (_m, sp: string, ch: string) => sp + ch.toLocaleUpperCase(ctx.locale));
    case 'firstcap': {
      const t = text();
      return t.charAt(0).toLocaleUpperCase(ctx.locale) + t.slice(1).toLocaleLowerCase(ctx.locale);
    }
    default: // MERGEFORMAT, CHARFORMAT and unknown formats leave the value untouched
      return value;
  }
}

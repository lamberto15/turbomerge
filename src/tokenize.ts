/**
 * Splits a Word field instruction into tokens, POSIX-shell style:
 * whitespace separates tokens, single or double quotes group text, quotes are removed,
 * adjacent quoted/unquoted text is concatenated and `""` yields an empty token.
 * There are no escape characters (Word paths use backslashes).
 *
 * Unlike Python's shlex (which throws on an unbalanced quote), an unterminated quote is
 * closed at the end of the instruction and reported through `onWarning`.
 */
export function tokenizeInstruction(instr: string, onWarning?: (message: string) => void): string[] {
  const tokens: string[] = [];
  let token = '';
  let inToken = false;
  let quote: string | null = null;

  for (const ch of instr) {
    if (quote !== null) {
      if (ch === quote) quote = null;
      else token += ch;
    } else if (ch === ' ' || ch === '\t' || ch === '\r' || ch === '\n') {
      if (inToken) {
        tokens.push(token);
        token = '';
        inToken = false;
      }
    } else if (ch === '"' || ch === "'") {
      quote = ch;
      inToken = true;
    } else {
      token += ch;
      inToken = true;
    }
  }
  if (quote !== null) onWarning?.(`Unbalanced quote in field instruction: <${instr.trim()}>`);
  if (inToken) tokens.push(token);
  return tokens;
}

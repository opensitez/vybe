// Generic printf-style formatter — bundled as `__vybe_sprintf`.
//
// Universal polyfill consumed by every language that needs C-style
// formatting (PHP `sprintf`/`printf`, Python `%` operator + `str.format`,
// Ruby `sprintf`/`format`/`%`, Pascal `Format`, Java/.NET `String.Format`
// where it routes to printf-style). Written in JS because the JS
// compiler is the most mature in vybex; the resulting bytecode chunk
// is language-agnostic.
//
// Coverage matches C99 + PHP extensions:
//   Conversions: %s %d %i %u %f %F %e %E %x %X %o %b %c %%
//   Flags: - (left-align) + (sign on positive) 0 (zero-pad)
//          ' ' (space for positive) # (alt form: 0x/0X/0/+ prefix)
//          'x (PHP custom pad character)
//   Width:  %05d %10s %-15s
//   Precision: %.3f %.5s %10.3f
//
// Built atop ECMA-262 §22.1 String / §21.1 Number primitives only —
// no `vybe:*` namespace, no host fns specific to sprintf.

function sprintf(fmt, ...args) {
    let out = "";
    let argIdx = 0;
    let i = 0;
    const len = fmt.length;
    while (i < len) {
        const c = fmt.charAt(i);
        if (c !== "%") {
            out += c;
            i++;
            continue;
        }
        i++;
        // Parse flags.
        let flagLeft = false, flagSign = false, flagZero = false;
        let flagSpace = false, flagAlt = false;
        while (i < len) {
            const f = fmt.charAt(i);
            if (f === "-") { flagLeft = true; i++; }
            else if (f === "+") { flagSign = true; i++; }
            else if (f === "0") { flagZero = true; i++; }
            else if (f === " ") { flagSpace = true; i++; }
            else if (f === "#") { flagAlt = true; i++; }
            else { break; }
        }
        // PHP custom pad: 'x.
        let padChar = null;
        if (i < len && fmt.charAt(i) === "'") {
            i++;
            if (i < len) { padChar = fmt.charAt(i); i++; }
        }
        // Width.
        let width = 0;
        while (i < len && fmt.charAt(i) >= "0" && fmt.charAt(i) <= "9") {
            width = width * 10 + (fmt.charCodeAt(i) - 48);
            i++;
        }
        // Precision.
        let precision = -1;
        if (i < len && fmt.charAt(i) === ".") {
            i++;
            precision = 0;
            while (i < len && fmt.charAt(i) >= "0" && fmt.charAt(i) <= "9") {
                precision = precision * 10 + (fmt.charCodeAt(i) - 48);
                i++;
            }
        }
        // Conversion specifier.
        if (i >= len) { out += "%"; break; }
        const conv = fmt.charAt(i);
        i++;
        if (conv === "%") { out += "%"; continue; }
        const arg = args[argIdx];
        argIdx++;
        let raw = "";
        if (conv === "s") {
            raw = String(arg);
            if (precision >= 0 && raw.length > precision) raw = raw.slice(0, precision);
        } else if (conv === "d" || conv === "i") {
            const n = Math.trunc(Number(arg));
            const neg = n < 0;
            let body = neg ? String(-n) : String(n);
            if (precision >= 0 && body.length < precision) {
                body = "0".repeat(precision - body.length) + body;
            }
            if (neg) raw = "-" + body;
            else if (flagSign) raw = "+" + body;
            else if (flagSpace) raw = " " + body;
            else raw = body;
        } else if (conv === "u") {
            const n = Number(arg);
            raw = String(n < 0 ? n + 0x100000000 : n);
        } else if (conv === "f" || conv === "F") {
            const n = Number(arg);
            const p = precision >= 0 ? precision : 6;
            let body = n.toFixed(p);
            if (n >= 0) {
                if (flagSign) body = "+" + body;
                else if (flagSpace) body = " " + body;
            }
            raw = body;
        } else if (conv === "e" || conv === "E") {
            const n = Number(arg);
            const p = precision >= 0 ? precision : 6;
            let body = n.toExponential(p);
            if (conv === "E") body = body.toUpperCase();
            raw = body;
        } else if (conv === "x") {
            const n = Number(arg);
            const u = n < 0 ? n + 0x100000000 : n;
            raw = u.toString(16);
            if (flagAlt) raw = "0x" + raw;
        } else if (conv === "X") {
            const n = Number(arg);
            const u = n < 0 ? n + 0x100000000 : n;
            raw = u.toString(16).toUpperCase();
            if (flagAlt) raw = "0X" + raw;
        } else if (conv === "o") {
            const n = Number(arg);
            const u = n < 0 ? n + 0x100000000 : n;
            raw = u.toString(8);
            if (flagAlt) raw = "0" + raw;
        } else if (conv === "b") {
            const n = Number(arg);
            const u = n < 0 ? n + 0x100000000 : n;
            raw = u.toString(2);
        } else if (conv === "c") {
            raw = String.fromCharCode(Number(arg));
        } else {
            // Unknown conversion: emit literally, return the arg slot.
            argIdx--;
            raw = "%" + conv;
        }
        // Apply width.
        if (raw.length < width) {
            const padLen = width - raw.length;
            let pc;
            if (padChar !== null) { pc = padChar; }
            else if (flagZero && !flagLeft) { pc = "0"; }
            else { pc = " "; }
            const pad = pc.repeat(padLen);
            if (flagLeft) {
                out += raw + pad;
            } else if (pc === "0" && (raw.charAt(0) === "-" || raw.charAt(0) === "+")) {
                // Zero-padding numbers must sit after the sign.
                out += raw.charAt(0) + pad + raw.slice(1);
            } else {
                out += pad + raw;
            }
        } else {
            out += raw;
        }
    }
    return out;
}

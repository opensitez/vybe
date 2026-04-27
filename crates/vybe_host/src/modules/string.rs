use std::sync::{Arc, Mutex};
use vybe_bytecode::{VM, Value, HostContext};

pub fn register(vm: &mut VM) {
    // ECMA-262 §22.1 String.prototype methods retired here:
    //   slice/indexOf/includes/split/replace/replaceAll/trimStart/trimEnd/
    //   search/match/concat/at/padStart/padEnd/startsWith/endsWith/charAt/
    //   substring/charCodeAt/fromCharCode/repeat
    // — full coverage now lives in `ecma:string` + `ecma:regexp` (search/match).
    // Callers compile to those addresses directly.

    // --- VB-compatible string functions (available to all languages) ---

    // left(str, n) → first n characters
    vm.register_host_fn("vybe:string", "left", Box::new(|_ctx, a| {
        let st = s(a, 0);
        let n = f(a, 1) as usize;
        let end = n.min(st.len());
        Value::String(Arc::from(&st[..end]))
    }));

    // `vybe:string.right` retired — VB `Right(s, n)` compiles to direct
    // opcodes via the `right` intrinsic in
    // `crates/vybex/src/compiler/mod.rs::emit_intrinsic` (STR_LENGTH +
    // I32_SUB + STR_SUBSTRING). No host call.

    // mid(str, start, length?) → substring (1-based start like VB)
    vm.register_host_fn("vybe:string", "mid", Box::new(|_ctx, a| {
        let st = s(a, 0);
        let start = ((f(a, 1) as usize).saturating_sub(1)).min(st.len());
        let end = if a.len() > 2 {
            (start + f(a, 2) as usize).min(st.len())
        } else {
            st.len()
        };
        Value::String(Arc::from(&st[start..end]))
    }));

    // instr(str, search) → 1-based position, 0 if not found
    // instr(start, str, search) → 1-based position starting from start
    vm.register_host_fn("vybe:string", "instr", Box::new(|_ctx, a| {
        if a.len() >= 3 {
            // instr(start, str, search)
            let start = (f(a, 0) as usize).saturating_sub(1);
            let st = s(a, 1);
            let search = s(a, 2);
            match st[start..].find(&search) {
                Some(idx) => Value::F64((start + idx + 1) as f64),
                None => Value::F64(0.0),
            }
        } else {
            // instr(str, search)
            let st = s(a, 0);
            let search = s(a, 1);
            match st.find(&search) {
                Some(idx) => Value::F64((idx + 1) as f64),
                None => Value::F64(0.0),
            }
        }
    }));

    // string(n, char) → string of n copies of char
    vm.register_host_fn("vybe:string", "stringRepeat", Box::new(|_ctx, a| {
        let n = f(a, 0) as usize;
        let ch = s(a, 1);
        let c = ch.chars().next().unwrap_or(' ');
        Value::String(Arc::from(c.to_string().repeat(n).as_str()))
    }));

    // instrrev(str, search) → 1-based position of LAST occurrence, 0 if not found
    vm.register_host_fn("vybe:string", "instrrev", Box::new(|_ctx, a| {
        let st = s(a, 0);
        let search = s(a, 1);
        match st.rfind(&search) {
            Some(idx) => Value::F64((idx + 1) as f64),
            None => Value::F64(0.0),
        }
    }));

    // format(value, formatStr) → formatted string
    // Supports both VB6 Format(value, spec) and .NET String.Format("{0}...", args...)
    vm.register_host_fn("vybe:string", "format", Box::new(|_ctx, a| {
        let first = s(a, 0);
        // Detect .NET composite format: first arg is a string containing {0}, {1}, etc.
        if first.contains("{0}") || first.contains("{1}") || first.contains("{2}") {
            let mut result = first.clone();
            for (i, arg) in a[1..].iter().enumerate() {
                let placeholder = format!("{{{}}}", i);
                result = result.replace(&placeholder, &format!("{}", arg));
            }
            Value::String(Arc::from(result.as_str()))
        } else {
            // VB6 Format(value, formatSpec)
            let val = f(a, 0);
            let fmt = s(a, 1).to_lowercase();
            let result = match fmt.as_str() {
                "0" | "0.0" | "#.#" | "fixed" => format!("{:.1}", val),
                "0.00" | "#.##" | "standard" => format!("{:.2}", val),
                "percent" => format!("{:.2}%", val * 100.0),
                "currency" => format!("${:.2}", val),
                "scientific" => format!("{:e}", val),
                "yes/no" => if val != 0.0 { "Yes".into() } else { "No".into() },
                "true/false" => if val != 0.0 { "True".into() } else { "False".into() },
                "on/off" => if val != 0.0 { "On".into() } else { "Off".into() },
                _ => format!("{}", val),
            };
            Value::String(Arc::from(result.as_str()))
        }
    }));

    // lset(str, length) → left-align in field
    vm.register_host_fn("vybe:string", "lset", Box::new(|_ctx, a| {
        let st = s(a, 0);
        let len = f(a, 1) as usize;
        Value::String(Arc::from(format!("{:<width$}", st, width = len).as_str()))
    }));

    // rset(str, length) → right-align in field
    vm.register_host_fn("vybe:string", "rset", Box::new(|_ctx, a| {
        let st = s(a, 0);
        let len = f(a, 1) as usize;
        Value::String(Arc::from(format!("{:>width$}", st, width = len).as_str()))
    }));

    // filter(arr, match, include?) → filtered array
    vm.register_host_fn("vybe:string", "filter", Box::new(|_ctx, a| {
        use std::sync::Mutex;
        use vybe_bytecode::value::{Object, ObjectKind};
        let match_str = s(a, 1);
        let include = if a.len() > 2 { f(a, 2) != 0.0 } else { true };
        let mut results = Vec::new();
        if let Some(Value::Object(obj)) = a.first() {
            let o = obj.lock().unwrap();
            if let ObjectKind::Array(ref elems) = o.kind {
                for elem in elems {
                    let es = format!("{}", elem);
                    let contains = es.contains(&match_str);
                    if (include && contains) || (!include && !contains) {
                        results.push(elem.clone());
                    }
                }
            }
        }
        Value::Object(Arc::new(Mutex::new(Object::new_array(results))))
    }));

    // count(str, sub) → number of non-overlapping occurrences
    vm.register_host_fn("vybe:string", "count", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let haystack = args.first().map(|v| format!("{}", v)).unwrap_or_default();
        let needle = args.get(1).map(|v| format!("{}", v)).unwrap_or_default();
        if needle.is_empty() {
            return Value::I32(0);
        }
        Value::I32(haystack.matches(&needle).count() as i32)
    }));

    // padStart(str, width) — zero-fill / right-justify
    vm.register_host_fn("vybe:string", "padStart", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let s = args.first().map(|v| format!("{}", v)).unwrap_or_default();
        let width = args.get(1).map(|v| v.as_f64() as usize).unwrap_or(0);
        let fill = args.get(2).map(|v| format!("{}", v)).unwrap_or_else(|| " ".to_string());
        let fill_char = fill.chars().next().unwrap_or(' ');
        if s.len() >= width {
            Value::String(Arc::from(s))
        } else {
            let padding: String = std::iter::repeat(fill_char).take(width - s.len()).collect();
            Value::String(Arc::from(format!("{}{}", padding, s)))
        }
    }));

    // sprintf — moved to the polyglot stdlib polyfill at
    // `crates/vybex/src/emitter/polyfills/sprintf.js`. Compiled once at
    // vybex build time via `build_polyfill(sprintf.js, "js", "sprintf")`,
    // bundled into every program as `__vybe_sprintf`. PHP / Python /
    // Ruby / Pascal sprintf-style callers compile to `stdlib:sprintf`.

    // ── phpIncrement(val) / phpDecrement(val) — PHP `++` / `--` ───────
    //
    // PHP defines `$s++` for strings as Perl-style character-carry
    // increment (not C-style string→number coercion):
    //   "aa"++ == "ab",  "az"++ == "ba",  "zz"++ == "aaa"
    //   "A9"++ == "B0",  "Z9"++ == "AA0"
    //   "2026-03-25"++ == "2026-03-26"   (carry stops at '-', which is
    //                                     not alphanumeric)
    // For non-string inputs, PHP numeric coercion applies and the
    // result is number + 1.
    //
    // These helpers are invoked via the PHP walker's AST rewrite of
    // `$x++` → `$x = __php_increment($x)` — the vybex compiler never
    // sees PHP-specific increment semantics.
    vm.register_host_fn("vybe:string", "phpIncrement", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        match args.first() {
            Some(Value::String(s)) => {
                Value::String(Arc::from(php_increment_string(s.as_ref()).as_str()))
            }
            Some(Value::Null) => Value::I32(1),
            Some(v) => Value::F64(v.as_f64() + 1.0),
            None => Value::Null,
        }
    }));

    // PHP `--` on strings is a no-op per spec (PHP manual:
    // "decrementing null values has no effect, but incrementing them
    // results in 1. decrementing string values […] has no effect").
    vm.register_host_fn("vybe:string", "phpDecrement", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        match args.first() {
            Some(Value::String(s)) => Value::String(Arc::clone(s)),
            Some(Value::Null) => Value::Null,
            Some(v) => Value::F64(v.as_f64() - 1.0),
            None => Value::Null,
        }
    }));

    // `vybe:string.substr` retired — PHP `substr($s, $start, $length?)`
    // compiles to direct opcodes via the `php_substr` intrinsic in
    // `crates/vybex/src/compiler/mod.rs::emit_intrinsic`, which composes
    // ECMA `substring(s, start, start + length)` semantics. No host call.
    // (PHP's negative-start / negative-length nuances are not yet
    // covered by the intrinsic — the common case `substr($s, $i, $n)`
    // with non-negative args matches ECMA substring exactly.)
}

// ── sprintf implementation ───────────────────────────────────────────
//
// State machine over format directives. PHP and C printf share the
// same directive grammar; PHP adds `%b` (binary) which we support.
fn sprintf(fmt: &str, args: &[Value]) -> String {
    let mut out = String::with_capacity(fmt.len());
    let mut chars = fmt.chars().peekable();
    let mut arg_idx = 0usize;
    while let Some(c) = chars.next() {
        if c != '%' {
            out.push(c);
            continue;
        }
        // Parse: flags [-+0 #], width (digits or *), .precision, conversion.
        let mut flag_left = false;
        let mut flag_sign = false;
        let mut flag_zero = false;
        let mut flag_space = false;
        let mut flag_alt = false;
        loop {
            match chars.peek() {
                Some('-') => { flag_left = true; chars.next(); }
                Some('+') => { flag_sign = true; chars.next(); }
                Some('0') => { flag_zero = true; chars.next(); }
                Some(' ') => { flag_space = true; chars.next(); }
                Some('#') => { flag_alt = true; chars.next(); }
                _ => break,
            }
        }
        // PHP-style padding char: `'x` selects `x` as the pad character.
        let mut pad_char: Option<char> = None;
        if chars.peek() == Some(&'\'') {
            chars.next();
            pad_char = chars.next();
        }
        let mut width: usize = 0;
        while let Some(&d) = chars.peek() {
            if d.is_ascii_digit() {
                width = width * 10 + (d as usize - '0' as usize);
                chars.next();
            } else { break; }
        }
        let mut precision: Option<usize> = None;
        if chars.peek() == Some(&'.') {
            chars.next();
            let mut p = 0usize;
            while let Some(&d) = chars.peek() {
                if d.is_ascii_digit() {
                    p = p * 10 + (d as usize - '0' as usize);
                    chars.next();
                } else { break; }
            }
            precision = Some(p);
        }
        let conv = chars.next().unwrap_or('%');
        if conv == '%' {
            out.push('%');
            continue;
        }
        let arg = args.get(arg_idx);
        arg_idx += 1;
        let raw = match conv {
            's' => {
                let mut s = arg.map(|v| format!("{}", v)).unwrap_or_default();
                if let Some(p) = precision { s.truncate(p); }
                s
            }
            'd' | 'i' => {
                let n = arg.map(|v| v.as_f64() as i64).unwrap_or(0);
                let mut s = if n < 0 {
                    format!("-{}", n.unsigned_abs())
                } else if flag_sign {
                    format!("+{}", n)
                } else if flag_space {
                    format!(" {}", n)
                } else {
                    format!("{}", n)
                };
                if let Some(p) = precision {
                    // Precision for d: minimum digits, zero-padded
                    let (sign, body) = if s.starts_with('-') || s.starts_with('+') || s.starts_with(' ') {
                        (s.chars().next().unwrap().to_string(), s[1..].to_string())
                    } else {
                        (String::new(), s.clone())
                    };
                    if body.len() < p {
                        let pad: String = std::iter::repeat('0').take(p - body.len()).collect();
                        s = format!("{}{}{}", sign, pad, body);
                    }
                }
                s
            }
            'u' => {
                let n = arg.map(|v| v.as_f64() as u64).unwrap_or(0);
                format!("{}", n)
            }
            'f' | 'F' => {
                let n = arg.map(|v| v.as_f64()).unwrap_or(0.0);
                let p = precision.unwrap_or(6);
                let s = format!("{:.*}", p, n);
                if flag_sign && n >= 0.0 { format!("+{}", s) }
                else if flag_space && n >= 0.0 { format!(" {}", s) }
                else { s }
            }
            'e' | 'E' => {
                let n = arg.map(|v| v.as_f64()).unwrap_or(0.0);
                let p = precision.unwrap_or(6);
                let s = format!("{:.*e}", p, n);
                if conv == 'E' { s.to_uppercase() } else { s }
            }
            'x' => {
                let n = arg.map(|v| v.as_f64() as u64).unwrap_or(0);
                let s = format!("{:x}", n);
                if flag_alt { format!("0x{}", s) } else { s }
            }
            'X' => {
                let n = arg.map(|v| v.as_f64() as u64).unwrap_or(0);
                let s = format!("{:X}", n);
                if flag_alt { format!("0X{}", s) } else { s }
            }
            'o' => {
                let n = arg.map(|v| v.as_f64() as u64).unwrap_or(0);
                let s = format!("{:o}", n);
                if flag_alt { format!("0{}", s) } else { s }
            }
            'b' => {
                let n = arg.map(|v| v.as_f64() as u64).unwrap_or(0);
                format!("{:b}", n)
            }
            'c' => {
                let code = arg.map(|v| v.as_f64() as u32).unwrap_or(0);
                char::from_u32(code).map(|c| c.to_string()).unwrap_or_default()
            }
            _ => {
                // Unknown directive — emit as literal so malformed templates
                // don't silently eat arguments.
                arg_idx -= 1;
                format!("%{}", conv)
            }
        };
        // Apply width.
        if raw.len() < width {
            let pad_len = width - raw.len();
            let pc = pad_char.unwrap_or(if flag_zero && !flag_left { '0' } else { ' ' });
            let pad: String = std::iter::repeat(pc).take(pad_len).collect();
            if flag_left {
                out.push_str(&raw);
                out.push_str(&pad);
            } else {
                // Zero-padding for numeric types must sit after the sign.
                if pc == '0' && (raw.starts_with('-') || raw.starts_with('+')) {
                    out.push(raw.chars().next().unwrap());
                    out.push_str(&pad);
                    out.push_str(&raw[1..]);
                } else {
                    out.push_str(&pad);
                    out.push_str(&raw);
                }
            }
        } else {
            out.push_str(&raw);
        }
    }
    out
}

fn s(args: &[Value], idx: usize) -> String {
    args.get(idx).map(|v| format!("{}", v)).unwrap_or_default()
}
fn f(args: &[Value], idx: usize) -> f64 {
    args.get(idx).map(|v| v.as_f64()).unwrap_or(0.0)
}
fn norm(idx: i64, len: i64) -> usize {
    if idx < 0 { (len + idx).max(0) as usize } else { idx.min(len) as usize }
}

/// Perl-style alphanumeric string increment. Mirrors PHP's
/// `increment_function` in ext/standard/incrementing.c — non-alphanumeric
/// characters halt the carry, and a carry out of the high end prepends
/// '1' (digit run) or 'a'/'A' (letter run) at the start of the run.
///
/// Empty string increments to `"1"` per PHP.
fn php_increment_string(s: &str) -> String {
    if s.is_empty() {
        return "1".to_string();
    }
    let mut bytes: Vec<u8> = s.as_bytes().to_vec();
    let mut i = bytes.len();
    let mut last_carry_char: Option<u8> = None;
    while i > 0 {
        i -= 1;
        let c = bytes[i];
        match c {
            b'0'..=b'9' => {
                if c == b'9' {
                    bytes[i] = b'0';
                    last_carry_char = Some(b'1');
                    // continue carrying left
                } else {
                    bytes[i] = c + 1;
                    return String::from_utf8(bytes).unwrap_or_else(|_| s.to_string());
                }
            }
            b'a'..=b'z' => {
                if c == b'z' {
                    bytes[i] = b'a';
                    last_carry_char = Some(b'a');
                } else {
                    bytes[i] = c + 1;
                    return String::from_utf8(bytes).unwrap_or_else(|_| s.to_string());
                }
            }
            b'A'..=b'Z' => {
                if c == b'Z' {
                    bytes[i] = b'A';
                    last_carry_char = Some(b'A');
                } else {
                    bytes[i] = c + 1;
                    return String::from_utf8(bytes).unwrap_or_else(|_| s.to_string());
                }
            }
            _ => {
                // Non-alphanumeric character: carry propagation halts.
                // If we had a pending carry, insert it immediately to the
                // right of this char (i.e. at the head of the run we were
                // incrementing). If there was no carry pending (we never
                // entered the run), return the string unchanged.
                if let Some(ch) = last_carry_char {
                    bytes.insert(i + 1, ch);
                }
                return String::from_utf8(bytes).unwrap_or_else(|_| s.to_string());
            }
        }
    }
    // Carried out of the left end — prepend the carry char.
    if let Some(ch) = last_carry_char {
        let mut out = Vec::with_capacity(bytes.len() + 1);
        out.push(ch);
        out.extend_from_slice(&bytes);
        return String::from_utf8(out).unwrap_or_else(|_| s.to_string());
    }
    String::from_utf8(bytes).unwrap_or_else(|_| s.to_string())
}

#[cfg(test)]
mod tests {
    use super::php_increment_string;
    #[test]
    fn php_inc_basic() {
        assert_eq!(php_increment_string(""), "1");
        assert_eq!(php_increment_string("a"), "b");
        assert_eq!(php_increment_string("z"), "aa");
        assert_eq!(php_increment_string("Az"), "Ba");
        assert_eq!(php_increment_string("zz"), "aaa");
        assert_eq!(php_increment_string("Zz"), "AAa");
        assert_eq!(php_increment_string("a9"), "b0");
        assert_eq!(php_increment_string("9"), "10");
        assert_eq!(php_increment_string("99"), "100");
    }
    #[test]
    fn php_inc_date_string() {
        // The hijridate loop case.
        assert_eq!(php_increment_string("2026-03-25"), "2026-03-26");
        assert_eq!(php_increment_string("2026-03-29"), "2026-03-30");
        assert_eq!(php_increment_string("2026-03-31"), "2026-03-32");
    }
    #[test]
    fn php_inc_carry_halt_on_nonalnum() {
        // Carry halts at '-'. Pending carry digit gets inserted to the
        // right of the non-alnum so "a-9" -> "a-10" (9 wraps to 0 with
        // carry, halts at '-', inserts '1').
        assert_eq!(php_increment_string("a-9"), "a-10");
    }
}

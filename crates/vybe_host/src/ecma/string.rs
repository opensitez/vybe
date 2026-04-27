//! `ecma:string` — ECMA-262 §22.1 String.
//!
//! Canonical JS-runtime string surface. Vybe-emitted .wasm calls
//! into these for `String.prototype.*` operations (the JS profile
//! routes `"foo".toUpperCase()` style calls here).
//!
//! Where the merged WebAssembly CG `wasm:js-string` proposal
//! covers an op (`length`, `concat`, `charCodeAt`, `substring`,
//! `equals`, `compare`, `fromCharCode`, `fromCodePoint`), this
//! module **does not reimplement it** — Vybe's compiler can emit
//! direct calls to `wasm:js-string` for those, and external CM
//! runtimes get a single source of truth. The functions registered
//! here cover the spec methods that aren't in the merged proposal:
//! casing, padding, trim, includes/indexOf/lastIndexOf, repeat,
//! replace/replaceAll, slice, split, startsWith/endsWith,
//! charAt/at/codePointAt, valueOf.
//!
//! Cross-language note: the legacy `vybe:string` module (.NET-shaped,
//! used by VB/C#/PHP/Python/etc. profiles) will eventually delegate
//! to these via thin forwarding shims so every language sees the
//! same JS-runtime semantics.

use std::sync::{Arc, Mutex};
use vybe_bytecode::value::{Object, ObjectKind};
use vybe_bytecode::{VM, Value};

fn s_arg(args: &[Value], idx: usize) -> String {
    match args.get(idx) {
        Some(Value::String(text)) => text.to_string(),
        Some(other) => format!("{}", other),
        None => String::new(),
    }
}

fn i32_arg(args: &[Value], idx: usize, default: i32) -> i32 {
    match args.get(idx) {
        Some(Value::I32(n)) => *n,
        Some(Value::F64(n)) => *n as i32,
        _ => default,
    }
}

fn s_val(text: &str) -> Value {
    Value::String(Arc::from(text))
}

/// Convert a possibly-negative ECMA-262 index into a clamped
/// non-negative position within `[0, len]`. Used by `slice` and
/// `at` which accept negative offsets.
fn clamp_signed(index: i32, len: usize) -> usize {
    let len_i = len as i32;
    let resolved = if index < 0 { (len_i + index).max(0) } else { index.min(len_i) };
    resolved.max(0) as usize
}

pub fn register(vm: &mut VM) {
    register_query_ops(vm);
    register_extract_ops(vm);
    register_casing_ops(vm);
    register_trim_ops(vm);
    register_pad_ops(vm);
    register_search_ops(vm);
    register_modify_ops(vm);
    register_split(vm);
    register_constructor_statics(vm);
    register_locale_compare(vm);
    register_normalize(vm);
    register_uri(vm);
    register_base64(vm);
    register_constructor(vm);
    register_adapters(vm);
}

// ── Adapter convenience methods ──────────────────────────────────
//
// Not in ECMA-262 §22.1 but ubiquitous across Python str /
// Ruby String / .NET — the predicates and case-conversions every
// language exposes. Live here as one-line Rust impls so the cross-
// language profile entries can target a single host fn surface (the
// `ecma:string` namespace) instead of the legacy `vybe:string`
// language-runtime catch-all. Same precedent as the ecma:array
// adapters (`clear`, `first`, `last`, `removeAt`, etc.).

fn register_adapters(vm: &mut VM) {
    // Python str.isdigit() — true iff non-empty and every char is
    // a Unicode decimal digit.
    vm.register_host_fn("ecma:string", "isdigit", Box::new(|_ctx, args| {
        let s = s_arg(args, 0);
        Value::Bool(!s.is_empty() && s.chars().all(|c| c.is_ascii_digit()))
    }));
    vm.register_host_fn("ecma:string", "isalpha", Box::new(|_ctx, args| {
        let s = s_arg(args, 0);
        Value::Bool(!s.is_empty() && s.chars().all(|c| c.is_alphabetic()))
    }));
    vm.register_host_fn("ecma:string", "isalnum", Box::new(|_ctx, args| {
        let s = s_arg(args, 0);
        Value::Bool(!s.is_empty() && s.chars().all(|c| c.is_alphanumeric()))
    }));
    vm.register_host_fn("ecma:string", "isspace", Box::new(|_ctx, args| {
        let s = s_arg(args, 0);
        Value::Bool(!s.is_empty() && s.chars().all(|c| c.is_whitespace()))
    }));
    vm.register_host_fn("ecma:string", "isupper", Box::new(|_ctx, args| {
        let s = s_arg(args, 0);
        let mut has_cased = false;
        for c in s.chars() {
            if c.is_uppercase() { has_cased = true; }
            else if c.is_lowercase() { return Value::Bool(false); }
        }
        Value::Bool(has_cased)
    }));
    vm.register_host_fn("ecma:string", "islower", Box::new(|_ctx, args| {
        let s = s_arg(args, 0);
        let mut has_cased = false;
        for c in s.chars() {
            if c.is_lowercase() { has_cased = true; }
            else if c.is_uppercase() { return Value::Bool(false); }
        }
        Value::Bool(has_cased)
    }));

    // Python str.title() / Ruby capitalize-each-word — uppercase the
    // first letter of each word, lowercase the rest.
    vm.register_host_fn("ecma:string", "title", Box::new(|_ctx, args| {
        let s = s_arg(args, 0);
        let mut out = String::with_capacity(s.len());
        let mut prev_alnum = false;
        for c in s.chars() {
            if c.is_alphanumeric() {
                if prev_alnum {
                    for lc in c.to_lowercase() { out.push(lc); }
                } else {
                    for uc in c.to_uppercase() { out.push(uc); }
                }
                prev_alnum = true;
            } else {
                out.push(c);
                prev_alnum = false;
            }
        }
        s_val(&out)
    }));

    // Python str.swapcase() — uppercase chars become lowercase and vice
    // versa; non-cased chars unchanged.
    vm.register_host_fn("ecma:string", "swapcase", Box::new(|_ctx, args| {
        let s = s_arg(args, 0);
        let mut out = String::with_capacity(s.len());
        for c in s.chars() {
            if c.is_uppercase() {
                for lc in c.to_lowercase() { out.push(lc); }
            } else if c.is_lowercase() {
                for uc in c.to_uppercase() { out.push(uc); }
            } else {
                out.push(c);
            }
        }
        s_val(&out)
    }));

    // Ruby `s.tr(from, to)` — translate chars in `from` to the parallel
    // char in `to`. Chars in `from` not in `to` are left unchanged.
    // Simplified: doesn't handle ranges (`a-z`) or negation (`^abc`) —
    // those need Ruby-specific intrinsics.
    vm.register_host_fn("ecma:string", "tr", Box::new(|_ctx, args| {
        let s = s_arg(args, 0);
        let from: Vec<char> = s_arg(args, 1).chars().collect();
        let to: Vec<char> = s_arg(args, 2).chars().collect();
        let out: String = s.chars().map(|c| {
            from.iter().position(|&fc| fc == c)
                .and_then(|i| to.get(i).copied())
                .unwrap_or(c)
        }).collect();
        s_val(&out)
    }));
}

// `String(v)` — §22.1.1.1 ToString(v) → §7.1.17.
//
// Per spec, for Objects: invoke ToPrimitive(v, "string") which
// dispatches to the object's `toString` method (preferring
// `Symbol.toPrimitive` if present, then `toString`, then `valueOf`).
// For primitives: use the Display impl which mirrors §7.1.17 Table 12.
//
// Mirrors the VM's `value_to_string` helper but invokes through
// HostContext.invoke so the dispatch works even when the constructor
// is called as a host fn from .NET / VB / etc. (where Convert.ToString
// expects method dispatch on objects rather than "[object Object]").
fn register_constructor(vm: &mut VM) {
    vm.register_host_fn("ecma:string", "String", Box::new(|ctx, args| {
        let v = args.first().cloned().unwrap_or(Value::Undefined);
        if let Value::Object(ref obj) = v {
            let to_str_fn = {
                let o = obj.lock().unwrap();
                o.properties.get("toString").cloned()
            };
            if let Some(fn_val) = to_str_fn {
                if matches!(fn_val, Value::Object(_)) {
                    let result = ctx.invoke(&fn_val, &[v.clone()]);
                    if !matches!(result, Value::Null | Value::Undefined) {
                        return s_val(&format!("{}", result));
                    }
                }
            }
        }
        s_val(&format!("{}", v))
    }));
}

// ── Query ops (length, char access) ───────────────────────────────

fn register_query_ops(vm: &mut VM) {
    // String.prototype.length — wasm:js-string already covers this
    // via a host fn. Re-register under ecma:string for callers that
    // want the canonical JS-runtime name.
    vm.register_host_fn("ecma:string", "length", Box::new(|_ctx, args| {
        Value::F64(s_arg(args, 0).chars().count() as f64)
    }));

    // String.prototype.charAt(pos)
    vm.register_host_fn("ecma:string", "charAt", Box::new(|_ctx, args| {
        let s = s_arg(args, 0);
        let pos = i32_arg(args, 1, 0);
        if pos < 0 { return s_val(""); }
        match s.chars().nth(pos as usize) {
            Some(ch) => s_val(&ch.to_string()),
            None => s_val(""),
        }
    }));

    // String.prototype.charCodeAt(pos)
    vm.register_host_fn("ecma:string", "charCodeAt", Box::new(|_ctx, args| {
        let s = s_arg(args, 0);
        let pos = i32_arg(args, 1, 0);
        if pos < 0 { return Value::F64(f64::NAN); }
        match s.chars().nth(pos as usize) {
            Some(ch) => Value::F64(ch as u32 as f64),
            None => Value::F64(f64::NAN),
        }
    }));

    // String.prototype.codePointAt(pos)
    vm.register_host_fn("ecma:string", "codePointAt", Box::new(|_ctx, args| {
        let s = s_arg(args, 0);
        let pos = i32_arg(args, 1, 0);
        if pos < 0 { return Value::Undefined; }
        match s.chars().nth(pos as usize) {
            Some(ch) => Value::F64(ch as u32 as f64),
            None => Value::Undefined,
        }
    }));

    // String.prototype.at(index) — ES2022; supports negative offsets.
    vm.register_host_fn("ecma:string", "at", Box::new(|_ctx, args| {
        let s = s_arg(args, 0);
        let index = i32_arg(args, 1, 0);
        let chars: Vec<char> = s.chars().collect();
        let len_i = chars.len() as i32;
        let resolved = if index < 0 { len_i + index } else { index };
        if resolved < 0 || resolved >= len_i { return Value::Undefined; }
        s_val(&chars[resolved as usize].to_string())
    }));

    // String.prototype.valueOf — returns the primitive string itself.
    vm.register_host_fn("ecma:string", "valueOf", Box::new(|_ctx, args| s_val(&s_arg(args, 0))));
}

// ── Extract ops (substring, slice, concat) ────────────────────────

fn register_extract_ops(vm: &mut VM) {
    // String.prototype.concat(...strings)
    vm.register_host_fn("ecma:string", "concat", Box::new(|_ctx, args| {
        let mut out = String::new();
        for a in args {
            match a {
                Value::String(text) => out.push_str(text),
                other => out.push_str(&format!("{}", other)),
            }
        }
        s_val(&out)
    }));

    // String.prototype.substring(indexStart, indexEnd?) — clamps
    // both args to [0, len], swaps if start > end.
    vm.register_host_fn("ecma:string", "substring", Box::new(|_ctx, args| {
        let s = s_arg(args, 0);
        let chars: Vec<char> = s.chars().collect();
        let len = chars.len();
        let start_raw = i32_arg(args, 1, 0).max(0) as usize;
        let start = start_raw.min(len);
        let end = if args.len() >= 3 {
            (i32_arg(args, 2, len as i32).max(0) as usize).min(len)
        } else {
            len
        };
        let (lo, hi) = if start <= end { (start, end) } else { (end, start) };
        s_val(&chars[lo..hi].iter().collect::<String>())
    }));

    // String.prototype.slice(start?, end?) — supports negative
    // offsets (count from end). Differs from substring: no swap on
    // out-of-order args (returns empty instead).
    vm.register_host_fn("ecma:string", "slice", Box::new(|_ctx, args| {
        let s = s_arg(args, 0);
        let chars: Vec<char> = s.chars().collect();
        let len = chars.len();
        let start = clamp_signed(i32_arg(args, 1, 0), len);
        let end = if args.len() >= 3 {
            clamp_signed(i32_arg(args, 2, len as i32), len)
        } else {
            len
        };
        if start >= end { return s_val(""); }
        s_val(&chars[start..end].iter().collect::<String>())
    }));
}

// ── Casing ops ────────────────────────────────────────────────────

fn register_casing_ops(vm: &mut VM) {
    vm.register_host_fn("ecma:string", "toUpperCase", Box::new(|_ctx, args| {
        s_val(&s_arg(args, 0).to_uppercase())
    }));
    vm.register_host_fn("ecma:string", "toLowerCase", Box::new(|_ctx, args| {
        s_val(&s_arg(args, 0).to_lowercase())
    }));
    // The locale-aware variants use the same Rust impl until
    // locale data is wired in; spec says the result is implementation
    // defined for non-locale data anyway.
    vm.register_host_fn("ecma:string", "toLocaleUpperCase", Box::new(|_ctx, args| {
        s_val(&s_arg(args, 0).to_uppercase())
    }));
    vm.register_host_fn("ecma:string", "toLocaleLowerCase", Box::new(|_ctx, args| {
        s_val(&s_arg(args, 0).to_lowercase())
    }));
}

// ── Trim ops ──────────────────────────────────────────────────────

fn register_trim_ops(vm: &mut VM) {
    vm.register_host_fn("ecma:string", "trim", Box::new(|_ctx, args| {
        s_val(s_arg(args, 0).trim())
    }));
    vm.register_host_fn("ecma:string", "trimStart", Box::new(|_ctx, args| {
        s_val(s_arg(args, 0).trim_start())
    }));
    vm.register_host_fn("ecma:string", "trimEnd", Box::new(|_ctx, args| {
        s_val(s_arg(args, 0).trim_end())
    }));
}

// ── Pad ops ───────────────────────────────────────────────────────

fn register_pad_ops(vm: &mut VM) {
    fn pad(args: &[Value], at_start: bool) -> Value {
        let s = s_arg(args, 0);
        let chars: Vec<char> = s.chars().collect();
        let target = i32_arg(args, 1, 0).max(0) as usize;
        if chars.len() >= target { return s_val(&s); }
        let pad_str = if args.len() >= 3 { s_arg(args, 2) } else { " ".to_string() };
        if pad_str.is_empty() { return s_val(&s); }
        let pad_chars: Vec<char> = pad_str.chars().collect();
        let needed = target - chars.len();
        let mut filler = String::with_capacity(needed);
        for i in 0..needed {
            filler.push(pad_chars[i % pad_chars.len()]);
        }
        let result = if at_start {
            format!("{}{}", filler, s)
        } else {
            format!("{}{}", s, filler)
        };
        s_val(&result)
    }
    vm.register_host_fn("ecma:string", "padStart", Box::new(|_ctx, args| pad(args, true)));
    vm.register_host_fn("ecma:string", "padEnd", Box::new(|_ctx, args| pad(args, false)));
}

// ── Search ops ────────────────────────────────────────────────────

fn register_search_ops(vm: &mut VM) {
    vm.register_host_fn("ecma:string", "includes", Box::new(|_ctx, args| {
        let s = s_arg(args, 0);
        let needle = s_arg(args, 1);
        let pos = i32_arg(args, 2, 0).max(0) as usize;
        if pos > s.len() { return Value::Bool(false); }
        Value::Bool(s[pos..].contains(&needle))
    }));

    vm.register_host_fn("ecma:string", "indexOf", Box::new(|_ctx, args| {
        let s = s_arg(args, 0);
        let needle = s_arg(args, 1);
        let pos = i32_arg(args, 2, 0).max(0) as usize;
        if pos > s.len() { return Value::F64(-1.0); }
        match s[pos..].find(&needle) {
            Some(idx) => Value::F64((pos + idx) as f64),
            None => Value::F64(-1.0),
        }
    }));

    vm.register_host_fn("ecma:string", "lastIndexOf", Box::new(|_ctx, args| {
        let s = s_arg(args, 0);
        let needle = s_arg(args, 1);
        match s.rfind(&needle) {
            Some(idx) => Value::F64(idx as f64),
            None => Value::F64(-1.0),
        }
    }));

    vm.register_host_fn("ecma:string", "startsWith", Box::new(|_ctx, args| {
        let s = s_arg(args, 0);
        let needle = s_arg(args, 1);
        let pos = i32_arg(args, 2, 0).max(0) as usize;
        if pos > s.len() { return Value::Bool(false); }
        Value::Bool(s[pos..].starts_with(&needle))
    }));

    vm.register_host_fn("ecma:string", "endsWith", Box::new(|_ctx, args| {
        let s = s_arg(args, 0);
        let needle = s_arg(args, 1);
        let end_pos = if args.len() >= 3 {
            (i32_arg(args, 2, s.len() as i32).max(0) as usize).min(s.len())
        } else {
            s.len()
        };
        Value::Bool(s[..end_pos].ends_with(&needle))
    }));
}

// ── Modify ops (repeat, replace, replaceAll) ──────────────────────

fn register_modify_ops(vm: &mut VM) {
    vm.register_host_fn("ecma:string", "repeat", Box::new(|_ctx, args| {
        let s = s_arg(args, 0);
        let n = i32_arg(args, 1, 0).max(0) as usize;
        s_val(&s.repeat(n))
    }));

    // ECMA-262 §22.1.3.18: replace with a string searchValue
    // replaces only the FIRST match. Rust's `str::replacen(.., 1)`
    // matches that.
    vm.register_host_fn("ecma:string", "replace", Box::new(|_ctx, args| {
        let s = s_arg(args, 0);
        let search = s_arg(args, 1);
        let replace = s_arg(args, 2);
        s_val(&s.replacen(&search, &replace, 1))
    }));

    // ECMA-262 §22.1.3.19: replaceAll replaces every occurrence.
    vm.register_host_fn("ecma:string", "replaceAll", Box::new(|_ctx, args| {
        let s = s_arg(args, 0);
        let search = s_arg(args, 1);
        let replace = s_arg(args, 2);
        s_val(&s.replace(&search, &replace))
    }));
}

// ── Split ─────────────────────────────────────────────────────────

fn register_split(vm: &mut VM) {
    vm.register_host_fn("ecma:string", "split", Box::new(|_ctx, args| {
        let s = s_arg(args, 0);
        let separator = match args.get(1) {
            Some(Value::String(text)) => Some(text.to_string()),
            None | Some(Value::Undefined) => None,
            Some(other) => Some(format!("{}", other)),
        };
        let limit: Option<usize> = match args.get(2) {
            Some(Value::F64(n)) if *n >= 0.0 => Some(*n as usize),
            Some(Value::I32(n)) if *n >= 0 => Some(*n as usize),
            _ => None,
        };

        let parts: Vec<Value> = match separator {
            None => vec![s_val(&s)],
            Some(sep) if sep.is_empty() => {
                // ECMA-262: empty string separator splits every char.
                s.chars().map(|c| s_val(&c.to_string())).collect()
            }
            Some(sep) => s.split(&sep).map(|piece| s_val(piece)).collect(),
        };
        let truncated: Vec<Value> = match limit {
            Some(n) => parts.into_iter().take(n).collect(),
            None => parts,
        };
        Value::Object(Arc::new(Mutex::new(Object::new_array(truncated))))
    }));
}

// ── String constructor statics ────────────────────────────────────

fn register_constructor_statics(vm: &mut VM) {
    vm.register_host_fn("ecma:string", "fromCharCode", Box::new(|_ctx, args| {
        let mut out = String::with_capacity(args.len());
        for arg in args {
            let code = match arg {
                Value::F64(n) => *n as u32,
                Value::I32(n) => *n as u32,
                _ => continue,
            };
            if let Some(ch) = char::from_u32(code) {
                out.push(ch);
            }
        }
        s_val(&out)
    }));

    vm.register_host_fn("ecma:string", "fromCodePoint", Box::new(|_ctx, args| {
        let mut out = String::with_capacity(args.len());
        for arg in args {
            let code = match arg {
                Value::F64(n) => *n as u32,
                Value::I32(n) => *n as u32,
                _ => continue,
            };
            if let Some(ch) = char::from_u32(code) {
                out.push(ch);
            }
        }
        s_val(&out)
    }));
}

// ── localeCompare (ECMA-262 §22.1.3.10) ───────────────────────────
//
// `a.localeCompare(b)` — locale-aware comparison. Returns -1 / 0 / 1
// per ECMA-402; the spec allows any negative / zero / positive return.
// Without an Intl.Collator implementation, this falls back to the
// codepoint ordering Rust's `str::cmp` produces — matches Node's
// default behaviour when Intl isn't available.

fn register_locale_compare(vm: &mut VM) {
    vm.register_host_fn("ecma:string", "localeCompare", Box::new(|_ctx, args| {
        let a = s_arg(args, 0);
        let b = s_arg(args, 1);
        Value::I32(match a.as_str().cmp(b.as_str()) {
            std::cmp::Ordering::Less => -1,
            std::cmp::Ordering::Equal => 0,
            std::cmp::Ordering::Greater => 1,
        })
    }));
}

// ── normalize (ECMA-262 §22.1.3.13) ───────────────────────────────
//
// `s.normalize(form?)` — Unicode normalization. Form is one of NFC /
// NFD / NFKC / NFKD; defaults to NFC. We don't ship a Unicode
// normalization library, so MVP returns the input unchanged for ASCII
// input and signals lossless passthrough for the common case. Full
// implementation requires the `unicode-normalization` crate.

fn register_normalize(vm: &mut VM) {
    vm.register_host_fn("ecma:string", "normalize", Box::new(|_ctx, args| {
        // ECMA-262 §22.1.3.13: returns the normalized String. Without
        // a normalization library, return the input unchanged — a
        // valid result for already-normalized strings (which include
        // all ASCII).
        s_val(&s_arg(args, 0))
    }));
}

// ── ECMA-262 §19.2.6 URI globals ───────────────────────────────────
//
// `encodeURI`, `decodeURI`, `encodeURIComponent`, `decodeURIComponent`
// are spec'd as global functions but they're string transforms — they
// belong on `ecma:string` for the purposes of the host-fn registry.

fn register_uri(vm: &mut VM) {
    // encodeURIComponent — encodes everything except the unreserved set
    // (ALPHA / DIGIT / `-` / `_` / `.` / `~` / `!` / `*` / `'` / `(` / `)`).
    // ECMA-262 §19.2.6.5.
    vm.register_host_fn("ecma:string", "encodeURIComponent", Box::new(|_ctx, args| {
        let s = s_arg(args, 0);
        let encoded: String = s.chars().flat_map(|c| {
            if c.is_ascii_alphanumeric() || "-_.!~*'()".contains(c) {
                c.to_string().chars().collect::<Vec<_>>()
            } else {
                let mut buf = [0u8; 4];
                let len = c.encode_utf8(&mut buf).len();
                buf[..len].iter()
                    .flat_map(|b| format!("%{:02X}", b).chars().collect::<Vec<_>>())
                    .collect()
            }
        }).collect();
        s_val(&encoded)
    }));

    // decodeURIComponent — reverses encodeURIComponent. Emits the raw
    // input on malformed escape sequences (ECMA spec throws URIError;
    // we degrade to passthrough until exception dispatch is wired).
    vm.register_host_fn("ecma:string", "decodeURIComponent", Box::new(|_ctx, args| {
        let s = s_arg(args, 0);
        let bytes = s.as_bytes();
        let mut out = Vec::with_capacity(bytes.len());
        let mut i = 0;
        while i < bytes.len() {
            if bytes[i] == b'%' && i + 2 < bytes.len() {
                if let Ok(byte) = u8::from_str_radix(std::str::from_utf8(&bytes[i+1..i+3]).unwrap_or(""), 16) {
                    out.push(byte);
                    i += 3;
                    continue;
                }
            }
            out.push(bytes[i]);
            i += 1;
        }
        s_val(&String::from_utf8_lossy(&out))
    }));

    // encodeURI — like encodeURIComponent but ALSO leaves URI-syntax
    // chars unencoded: `;` `,` `/` `?` `:` `@` `&` `=` `+` `$` `#`.
    // ECMA-262 §19.2.6.4.
    vm.register_host_fn("ecma:string", "encodeURI", Box::new(|_ctx, args| {
        let s = s_arg(args, 0);
        let encoded: String = s.chars().flat_map(|c| {
            if c.is_ascii_alphanumeric() || "-_.!~*'();,/?:@&=+$#".contains(c) {
                c.to_string().chars().collect::<Vec<_>>()
            } else {
                let mut buf = [0u8; 4];
                let len = c.encode_utf8(&mut buf).len();
                buf[..len].iter()
                    .flat_map(|b| format!("%{:02X}", b).chars().collect::<Vec<_>>())
                    .collect()
            }
        }).collect();
        s_val(&encoded)
    }));

    // decodeURI — reverses encodeURI. The spec defines a different
    // reserved set than decodeURIComponent (it preserves URI-syntax
    // chars even if they were percent-encoded), but for our MVP we
    // simply unescape every `%XX` — same behaviour as decodeURIComponent.
    vm.register_host_fn("ecma:string", "decodeURI", Box::new(|_ctx, args| {
        let s = s_arg(args, 0);
        let bytes = s.as_bytes();
        let mut out = Vec::with_capacity(bytes.len());
        let mut i = 0;
        while i < bytes.len() {
            if bytes[i] == b'%' && i + 2 < bytes.len() {
                if let Ok(byte) = u8::from_str_radix(std::str::from_utf8(&bytes[i+1..i+3]).unwrap_or(""), 16) {
                    out.push(byte);
                    i += 3;
                    continue;
                }
            }
            out.push(bytes[i]);
            i += 1;
        }
        s_val(&String::from_utf8_lossy(&out))
    }));
}

// ── WHATWG btoa / atob ────────────────────────────────────────────
//
// Not strictly ECMA-262 but JS exposes them as globals (HTML spec
// §8.3 — WindowOrWorkerGlobalScope). Same shape as the URI helpers
// (string in, string out), so they live here next to encodeURIComponent.

const BASE64_CHARS: &[u8] = b"ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/";

fn base64_encode(data: &[u8]) -> String {
    let mut result = String::with_capacity(((data.len() + 2) / 3) * 4);
    for chunk in data.chunks(3) {
        let b0 = chunk[0] as u32;
        let b1 = if chunk.len() > 1 { chunk[1] as u32 } else { 0 };
        let b2 = if chunk.len() > 2 { chunk[2] as u32 } else { 0 };
        let n = (b0 << 16) | (b1 << 8) | b2;
        result.push(BASE64_CHARS[((n >> 18) & 63) as usize] as char);
        result.push(BASE64_CHARS[((n >> 12) & 63) as usize] as char);
        if chunk.len() > 1 { result.push(BASE64_CHARS[((n >> 6) & 63) as usize] as char); } else { result.push('='); }
        if chunk.len() > 2 { result.push(BASE64_CHARS[(n & 63) as usize] as char); } else { result.push('='); }
    }
    result
}

fn base64_decode(s: &str) -> Option<Vec<u8>> {
    const DECODE: [u8; 128] = {
        let mut t = [255u8; 128];
        let mut i = 0u8;
        while i < 26 { t[(b'A' + i) as usize] = i; i += 1; }
        i = 0;
        while i < 26 { t[(b'a' + i) as usize] = 26 + i; i += 1; }
        i = 0;
        while i < 10 { t[(b'0' + i) as usize] = 52 + i; i += 1; }
        t[b'+' as usize] = 62;
        t[b'/' as usize] = 63;
        t
    };
    let bytes: Vec<u8> = s.bytes().filter(|&b| b != b'=' && b != b'\n' && b != b'\r').collect();
    let mut result = Vec::new();
    for chunk in bytes.chunks(4) {
        if chunk.len() < 2 { break; }
        let a = *DECODE.get(chunk[0] as usize)? as u32;
        let b = *DECODE.get(chunk[1] as usize)? as u32;
        if a == 255 || b == 255 { return None; }
        result.push(((a << 2) | (b >> 4)) as u8);
        if chunk.len() > 2 {
            let c = *DECODE.get(chunk[2] as usize)? as u32;
            if c == 255 { return None; }
            result.push((((b & 0xF) << 4) | (c >> 2)) as u8);
            if chunk.len() > 3 {
                let d = *DECODE.get(chunk[3] as usize)? as u32;
                if d == 255 { return None; }
                result.push((((c & 0x3) << 6) | d) as u8);
            }
        }
    }
    Some(result)
}

fn register_base64(vm: &mut VM) {
    // btoa — base64-encode the input string's BYTES (treating it as
    // Latin-1 per HTML spec; we use the UTF-8 bytes which matches
    // most modern usage).
    vm.register_host_fn("ecma:string", "btoa", Box::new(|_ctx, args| {
        let s = s_arg(args, 0);
        s_val(&base64_encode(s.as_bytes()))
    }));

    // atob — base64-decode and interpret bytes as a string. Returns
    // null on malformed input (HTML spec throws DOMException; we
    // degrade to null until exception dispatch is wired).
    vm.register_host_fn("ecma:string", "atob", Box::new(|_ctx, args| {
        let s = s_arg(args, 0);
        match base64_decode(&s) {
            Some(bytes) => s_val(&String::from_utf8_lossy(&bytes)),
            None => Value::Null,
        }
    }));
}

#[allow(dead_code)]
fn _force_object_use(_: ObjectKind) {}

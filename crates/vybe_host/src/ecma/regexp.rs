//! `ecma:regexp` — ECMA-262 §22.2 RegExp + the regex-taking
//! `String.prototype` methods (match, matchAll, search, replace,
//! replaceAll, split).
//!
//! Backed by the audited `regex` crate. Pattern + flags compose into a
//! single Rust regex string via the `(?<flags>)` inline modifier syntax,
//! so JS `new RegExp("hello", "i")` lowers to `(?i)hello`.
//!
//! JS flag → Rust inline modifier mapping (ECMA-262 §22.2.5.1):
//!   `i` → `i` (case-insensitive)
//!   `m` → `m` (multi-line ^/$)
//!   `s` → `s` (dotAll: . matches newline)
//!   `u` → no-op (Rust regex is always Unicode-aware)
//!   `g` → no inline equivalent — handled by caller (find_iter vs find)
//!   `y` → not supported by the `regex` crate (sticky flag); ignored
//!   `d` → not supported (match indices); ignored
//!
//! Construct shape:
//!   - `ObjectKind::Ordinary` with properties `source`, `flags`, `global`,
//!     `ignoreCase`, `multiline`, `dotAll`, `unicode`, `sticky`,
//!     `lastIndex`, `__type=RegExp`. The `__type` stamp lets
//!     `instanceof RegExp` work via the cross-language type registry.

use std::sync::{Arc, Mutex};
use vybe_bytecode::value::{Object, ObjectKind, Value};
use vybe_bytecode::{HostContext, VM};

const REGEXP_TYPE: &str = "RegExp";

fn s_arg(args: &[Value], idx: usize) -> String {
    match args.get(idx) {
        Some(Value::String(s)) => s.to_string(),
        Some(other) => format!("{}", other),
        None => String::new(),
    }
}

/// Extract `source` + `flags` from a RegExp object (or treat raw string
/// args as `(pattern, flags)`). Returns `(pattern, flags)` strings.
fn extract_pattern(args: &[Value], idx: usize) -> (String, String) {
    match args.get(idx) {
        Some(Value::Object(obj)) => {
            let o = obj.lock().unwrap();
            let src = o.properties.get("source")
                .map(|v| match v { Value::String(s) => s.to_string(), o => format!("{}", o) })
                .unwrap_or_default();
            let flags = o.properties.get("flags")
                .map(|v| match v { Value::String(s) => s.to_string(), o => format!("{}", o) })
                .unwrap_or_default();
            (src, flags)
        }
        Some(Value::String(s)) => (s.to_string(), String::new()),
        Some(other) => (format!("{}", other), String::new()),
        None => (String::new(), String::new()),
    }
}

/// Map JS flags into the Rust regex inline-modifier prefix.
/// Returns the compiled `regex::Regex` (or `None` if pattern is invalid).
fn compile(pattern: &str, flags: &str) -> Option<regex::Regex> {
    let mut inline = String::new();
    for c in flags.chars() {
        match c {
            'i' => inline.push('i'),
            'm' => inline.push('m'),
            's' => inline.push('s'),
            // 'u', 'g', 'y', 'd' have no Rust inline modifier; see header doc.
            _ => {}
        }
    }
    let full = if inline.is_empty() {
        pattern.to_string()
    } else {
        format!("(?{}){}", inline, pattern)
    };
    regex::Regex::new(&full).ok()
}

fn make_array(elements: Vec<Value>) -> Value {
    Value::Object(Arc::new(Mutex::new(Object::new_array(elements))))
}

fn s_val(s: &str) -> Value {
    Value::String(Arc::from(s))
}

pub fn register(vm: &mut VM) {
    register_constructor(vm);
    register_prototype(vm);
    register_string_methods(vm);
}

// ── Constructor ───────────────────────────────────────────────────────

fn register_constructor(vm: &mut VM) {
    // `new RegExp(pattern, flags?)` — ECMA-262 §22.2.4. Pattern can be a
    // string or another RegExp (in which case its source/flags are reused).
    vm.register_host_fn("ecma:regexp", "new",
        Box::new(|_ctx, args| {
            let (pattern, default_flags) = extract_pattern(args, 0);
            // Explicit flags arg overrides any flags inherited from a
            // RegExp first arg.
            let flags = match args.get(1) {
                Some(Value::String(s)) => s.to_string(),
                Some(Value::Undefined) | None => default_flags,
                Some(other) => format!("{}", other),
            };
            let mut obj = Object::new();
            obj.properties.insert("source".into(), s_val(&pattern));
            obj.properties.insert("flags".into(), s_val(&flags));
            obj.properties.insert("global".into(), Value::Bool(flags.contains('g')));
            obj.properties.insert("ignoreCase".into(), Value::Bool(flags.contains('i')));
            obj.properties.insert("multiline".into(), Value::Bool(flags.contains('m')));
            obj.properties.insert("dotAll".into(), Value::Bool(flags.contains('s')));
            obj.properties.insert("unicode".into(), Value::Bool(flags.contains('u')));
            obj.properties.insert("sticky".into(), Value::Bool(flags.contains('y')));
            obj.properties.insert("hasIndices".into(), Value::Bool(flags.contains('d')));
            obj.properties.insert("lastIndex".into(), Value::I32(0));
            // __type lets cross-language `instanceof RegExp` work via the
            // type registry; matches the pattern used by Map/Set/etc.
            obj.properties.insert("__type".into(), Value::String(Arc::from(REGEXP_TYPE)));
            Value::Object(Arc::new(Mutex::new(obj)))
        }));
}

// ── RegExp.prototype ─────────────────────────────────────────────────

fn register_prototype(vm: &mut VM) {
    // `regex.test(str)` — ECMA-262 §22.2.5.15. True iff pattern matches
    // anywhere in str. Receiver is `args[0]` per Component-Model
    // `[method]` convention.
    vm.register_host_fn("ecma:regexp", "test",
        Box::new(|_ctx, args| {
            let (pattern, flags) = extract_pattern(args, 0);
            let input = s_arg(args, 1);
            match compile(&pattern, &flags) {
                Some(re) => Value::Bool(re.is_match(&input)),
                None => Value::Bool(false),
            }
        }));

    // `regex.exec(str)` — ECMA-262 §22.2.5.2. Returns a match Array
    // `[full, g1, g2, ..., index, input, groups]` or null.
    //
    // Spec layout: the array's numeric elements are full + capture groups,
    // with `.index`, `.input`, and `.groups` set as own properties on the
    // array. We materialize all of these so `match[0]`, `match.index`,
    // and `match.groups.name` all work.
    vm.register_host_fn("ecma:regexp", "exec",
        Box::new(|_ctx, args| {
            let (pattern, flags) = extract_pattern(args, 0);
            let input = s_arg(args, 1);
            let re = match compile(&pattern, &flags) {
                Some(re) => re,
                None => return Value::Null,
            };
            let caps = match re.captures(&input) {
                Some(c) => c,
                None => return Value::Null,
            };
            // Numeric elements: full match + each capture group.
            let mut elems: Vec<Value> = Vec::with_capacity(caps.len());
            for i in 0..caps.len() {
                elems.push(match caps.get(i) {
                    Some(m) => s_val(m.as_str()),
                    None => Value::Undefined,
                });
            }
            let mut match_obj = Object::new_array(elems);
            // Spec sets `index`, `input`, `groups` as own properties on the
            // returned Array.
            let index = caps.get(0).map(|m| m.start() as i32).unwrap_or(0);
            match_obj.properties.insert("index".into(), Value::I32(index));
            match_obj.properties.insert("input".into(), s_val(&input));
            // Named groups
            let mut groups = Object::new();
            for name in re.capture_names().flatten() {
                let val = caps.name(name).map(|m| s_val(m.as_str())).unwrap_or(Value::Undefined);
                groups.properties.insert(name.to_string(), val);
            }
            match_obj.properties.insert("groups".into(),
                Value::Object(Arc::new(Mutex::new(groups))));
            Value::Object(Arc::new(Mutex::new(match_obj)))
        }));

    // `regex.toString()` — ECMA-262 §22.2.5.17. Returns "/source/flags".
    vm.register_host_fn("ecma:regexp", "toString",
        Box::new(|_ctx, args| {
            let (pattern, flags) = extract_pattern(args, 0);
            s_val(&format!("/{}/{}", pattern, flags))
        }));
}

// ── String.prototype regex methods ───────────────────────────────────
//
// These take a string receiver + RegExp argument. Live under
// `ecma:regexp` (rather than `ecma:string`) because the regex compiler
// is the load-bearing dependency — keeping all regex-using ops in one
// place makes flag handling consistent and lets the `regex` crate
// dependency stay scoped to this file.

fn register_string_methods(vm: &mut VM) {
    // `str.match(regex)` — §22.1.3.13. Without `g`: same as
    // `regex.exec(str)` (single match Array with groups). With `g`:
    // Array of full-match strings only (no groups).
    vm.register_host_fn("ecma:regexp", "match",
        Box::new(|_ctx, args| {
            let input = s_arg(args, 0);
            let (pattern, flags) = extract_pattern(args, 1);
            let re = match compile(&pattern, &flags) {
                Some(re) => re,
                None => return Value::Null,
            };
            if flags.contains('g') {
                // Global: array of full matches only, no groups, no index.
                let matches: Vec<Value> = re.find_iter(&input)
                    .map(|m| s_val(m.as_str()))
                    .collect();
                if matches.is_empty() {
                    Value::Null
                } else {
                    make_array(matches)
                }
            } else {
                // Non-global: same shape as exec.
                let caps = match re.captures(&input) {
                    Some(c) => c,
                    None => return Value::Null,
                };
                let mut elems: Vec<Value> = Vec::with_capacity(caps.len());
                for i in 0..caps.len() {
                    elems.push(match caps.get(i) {
                        Some(m) => s_val(m.as_str()),
                        None => Value::Undefined,
                    });
                }
                let mut match_obj = Object::new_array(elems);
                let index = caps.get(0).map(|m| m.start() as i32).unwrap_or(0);
                match_obj.properties.insert("index".into(), Value::I32(index));
                match_obj.properties.insert("input".into(), s_val(&input));
                let mut groups = Object::new();
                for name in re.capture_names().flatten() {
                    let val = caps.name(name).map(|m| s_val(m.as_str())).unwrap_or(Value::Undefined);
                    groups.properties.insert(name.to_string(), val);
                }
                match_obj.properties.insert("groups".into(),
                    Value::Object(Arc::new(Mutex::new(groups))));
                Value::Object(Arc::new(Mutex::new(match_obj)))
            }
        }));

    // `str.matchAll(regex)` — §22.1.3.14. Spec returns an iterator;
    // MVP returns an Array of match Arrays (each shaped like exec's
    // result). Iterator semantics layer on top once iterator protocol
    // dispatch lands.
    vm.register_host_fn("ecma:regexp", "matchAll",
        Box::new(|_ctx, args| {
            let input = s_arg(args, 0);
            let (pattern, flags) = extract_pattern(args, 1);
            let re = match compile(&pattern, &flags) {
                Some(re) => re,
                None => return make_array(Vec::new()),
            };
            let mut out = Vec::new();
            for caps in re.captures_iter(&input) {
                let mut elems: Vec<Value> = Vec::with_capacity(caps.len());
                for i in 0..caps.len() {
                    elems.push(match caps.get(i) {
                        Some(m) => s_val(m.as_str()),
                        None => Value::Undefined,
                    });
                }
                let mut match_obj = Object::new_array(elems);
                let index = caps.get(0).map(|m| m.start() as i32).unwrap_or(0);
                match_obj.properties.insert("index".into(), Value::I32(index));
                match_obj.properties.insert("input".into(), s_val(&input));
                out.push(Value::Object(Arc::new(Mutex::new(match_obj))));
            }
            make_array(out)
        }));

    // `str.search(regex)` — §22.1.3.16. Returns index of first match
    // or -1.
    vm.register_host_fn("ecma:regexp", "search",
        Box::new(|_ctx, args| {
            let input = s_arg(args, 0);
            let (pattern, flags) = extract_pattern(args, 1);
            match compile(&pattern, &flags) {
                Some(re) => match re.find(&input) {
                    Some(m) => Value::I32(m.start() as i32),
                    None => Value::I32(-1),
                },
                None => Value::I32(-1),
            }
        }));

    // `str.replace(regex, replacement)` — §22.1.3.18. Replaces first
    // match (or all if `g` flag is set, per spec). Replacement string
    // supports `$1`/`$2`/`$<name>` capture refs via the `regex` crate.
    vm.register_host_fn("ecma:regexp", "replace",
        Box::new(|_ctx, args| {
            let input = s_arg(args, 0);
            let (pattern, flags) = extract_pattern(args, 1);
            let replacement = s_arg(args, 2);
            let re = match compile(&pattern, &flags) {
                Some(re) => re,
                None => return s_val(&input),
            };
            // Spec: `g` flag → replace all; else replace first.
            let result = if flags.contains('g') {
                re.replace_all(&input, replacement.as_str()).into_owned()
            } else {
                re.replace(&input, replacement.as_str()).into_owned()
            };
            s_val(&result)
        }));

    // `str.replaceAll(regex, replacement)` — §22.1.3.19. With a RegExp,
    // requires the `g` flag (otherwise spec throws TypeError); we just
    // replace-all unconditionally for simplicity.
    vm.register_host_fn("ecma:regexp", "replaceAll",
        Box::new(|_ctx, args| {
            let input = s_arg(args, 0);
            let (pattern, flags) = extract_pattern(args, 1);
            let replacement = s_arg(args, 2);
            match compile(&pattern, &flags) {
                Some(re) => s_val(&re.replace_all(&input, replacement.as_str())),
                None => s_val(&input),
            }
        }));

    // `str.split(regex, limit?)` — §22.1.3.20. Splits on regex matches.
    vm.register_host_fn("ecma:regexp", "split",
        Box::new(|_ctx, args| {
            let input = s_arg(args, 0);
            let (pattern, flags) = extract_pattern(args, 1);
            let limit = args.get(2)
                .map(|v| v.as_i32())
                .filter(|n| *n > 0)
                .map(|n| n as usize);
            match compile(&pattern, &flags) {
                Some(re) => {
                    let parts: Vec<Value> = match limit {
                        Some(n) => re.splitn(&input, n).map(s_val).collect(),
                        None => re.split(&input).map(s_val).collect(),
                    };
                    make_array(parts)
                }
                None => make_array(vec![s_val(&input)]),
            }
        }));
}

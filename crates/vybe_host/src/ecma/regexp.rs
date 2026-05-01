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
use vybe_bytecode::value::{Object, Value};
use vybe_bytecode::VM;

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
        Some(Value::String(s)) => split_regex_literal(s.as_ref()),
        Some(other) => split_regex_literal(&format!("{}", other)),
        None => (String::new(), String::new()),
    }
}

/// Pull pattern + flags out of a string in `/pat/flags` shape — what the
/// JS walker emits when it encounters a regex literal `/\d+/g`. The
/// pattern may contain escaped slashes (`\/`); we split on the LAST
/// unescaped `/`. Plain strings (no leading `/`) pass through as the
/// pattern with empty flags.
fn split_regex_literal(s: &str) -> (String, String) {
    if !s.starts_with('/') {
        return (s.to_string(), String::new());
    }
    // Find the LAST `/` not preceded by an odd number of backslashes.
    let bytes = s.as_bytes();
    let mut last = None;
    for (i, &b) in bytes.iter().enumerate().skip(1) {
        if b == b'/' {
            let mut bs = 0;
            let mut k = i;
            while k > 0 && bytes[k - 1] == b'\\' { bs += 1; k -= 1; }
            if bs % 2 == 0 { last = Some(i); }
        }
    }
    match last {
        Some(end) if end > 0 => {
            let pattern = s[1..end].to_string();
            let flags = s[end + 1..].to_string();
            (pattern, flags)
        }
        _ => (s.to_string(), String::new()),
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
        Box::new(|_ctx, args| regexp_test(args)));

    // `regex.exec(str)` — ECMA-262 §22.2.5.2. Returns a match Array
    // `[full, g1, g2, ..., index, input, groups]` or null.
    //
    // Spec layout: the array's numeric elements are full + capture groups,
    // with `.index`, `.input`, and `.groups` set as own properties on the
    // array. We materialize all of these so `match[0]`, `match.index`,
    // and `match.groups.name` all work.
    vm.register_host_fn("ecma:regexp", "exec",
        Box::new(|_ctx, args| regexp_exec(args)));

    // `regex.toString()` — ECMA-262 §22.2.5.17. Returns "/source/flags".
    vm.register_host_fn("ecma:regexp", "toString",
        Box::new(|_ctx, args| regexp_to_string(args)));
}

pub fn dispatch_regexp_method(method: &str, args: &[Value]) -> Option<Value> {
    match method {
        "test" => Some(regexp_test(args)),
        "exec" => Some(regexp_exec(args)),
        "toString" => Some(regexp_to_string(args)),
        _ => None,
    }
}

fn regexp_test(args: &[Value]) -> Value {
    let (pattern, flags) = extract_pattern(args, 0);
    let input = s_arg(args, 1);
    match compile(&pattern, &flags) {
        Some(re) => Value::Bool(re.is_match(&input)),
        None => Value::Bool(false),
    }
}

fn regexp_exec(args: &[Value]) -> Value {
    let (pattern, flags) = extract_pattern(args, 0);
    let input = s_arg(args, 1);
    let re = match compile(&pattern, &flags) {
        Some(re) => re,
        None => return Value::Null,
    };
    let is_global_or_sticky = flags.contains('g') || flags.contains('y');
    let last_index = if is_global_or_sticky {
        args.first().and_then(|v| match v {
            Value::Object(obj) => obj.lock().unwrap().properties.get("lastIndex").map(|v| v.as_i32()),
            _ => None,
        }).unwrap_or(0).max(0) as usize
    } else { 0 };
    let search_start = last_index.min(input.len());
    let caps = match re.captures(&input[search_start..]) {
        Some(c) => c,
        None => {
            if is_global_or_sticky {
                if let Some(Value::Object(obj)) = args.first() {
                    obj.lock().unwrap().properties.insert("lastIndex".into(), Value::I32(0));
                }
            }
            return Value::Null;
        }
    };
    if is_global_or_sticky {
        let new_idx = caps.get(0).map(|m| (search_start + m.end()) as i32).unwrap_or(0);
        if let Some(Value::Object(obj)) = args.first() {
            obj.lock().unwrap().properties.insert("lastIndex".into(), Value::I32(new_idx));
        }
    }
    let mut elems: Vec<Value> = Vec::with_capacity(caps.len());
    for i in 0..caps.len() {
        elems.push(match caps.get(i) {
            Some(m) => s_val(m.as_str()),
            None => Value::Undefined,
        });
    }
    let mut match_obj = Object::new_array(elems);
    let index = caps.get(0).map(|m| (search_start + m.start()) as i32).unwrap_or(0);
    match_obj.properties.insert("index".into(), Value::I32(index));
    match_obj.properties.insert("input".into(), s_val(&input));
    let mut groups = Object::new();
    let mut group_order: Vec<Value> = Vec::new();
    for name in re.capture_names().flatten() {
        let val = caps.name(name).map(|m| s_val(m.as_str())).unwrap_or(Value::Undefined);
        groups.properties.insert(name.to_string(), val);
        group_order.push(s_val(name));
    }
    if !group_order.is_empty() {
        groups.properties.insert("__keys".into(),
            Value::Object(Arc::new(Mutex::new(Object::new_array(group_order)))));
    }
    match_obj.properties.insert("groups".into(),
        Value::Object(Arc::new(Mutex::new(groups))));
    Value::Object(Arc::new(Mutex::new(match_obj)))
}

fn regexp_to_string(args: &[Value]) -> Value {
    let (pattern, flags) = extract_pattern(args, 0);
    s_val(&format!("/{}/{}", pattern, flags))
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
                let mut group_order: Vec<Value> = Vec::new();
                for name in re.capture_names().flatten() {
                    let val = caps.name(name).map(|m| s_val(m.as_str())).unwrap_or(Value::Undefined);
                    groups.properties.insert(name.to_string(), val);
                    group_order.push(s_val(name));
                }
                if !group_order.is_empty() {
                    groups.properties.insert("__keys".into(),
                        Value::Object(Arc::new(Mutex::new(Object::new_array(group_order)))));
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
    // match (or all if `g` flag is set, per spec). Replacement is either
    // a string (with $1/$2/$<name> capture refs) or a function called
    // with (match, ...captures, offset, input). The function form needs
    // VM callback dispatch via `ctx.invoke`.
    vm.register_host_fn("ecma:regexp", "replace",
        Box::new(|ctx, args| {
            let input = s_arg(args, 0);
            let (pattern, flags) = extract_pattern(args, 1);
            let re = match compile(&pattern, &flags) {
                Some(re) => re,
                None => return s_val(&input),
            };
            let global = flags.contains('g');
            // Function-form replacement: invoke the callback per match
            // and substitute its return value. Spec §22.1.3.18 step 8.b.iii.
            let replacement_arg = args.get(2).cloned().unwrap_or(Value::Undefined);
            let is_callable = matches!(&replacement_arg, Value::Object(o)
                if matches!(o.lock().unwrap().kind,
                    vybe_bytecode::value::ObjectKind::Function(_)
                    | vybe_bytecode::value::ObjectKind::HostFunction(_)));
            if is_callable {
                let mut out = String::with_capacity(input.len());
                let mut last_end = 0;
                let captures_iter: Box<dyn Iterator<Item = regex::Captures>> = if global {
                    Box::new(re.captures_iter(&input))
                } else {
                    Box::new(re.captures(&input).into_iter())
                };
                for caps in captures_iter {
                    let m = match caps.get(0) { Some(m) => m, None => continue };
                    out.push_str(&input[last_end..m.start()]);
                    // Build callback args: match, ...groups, offset, input
                    let mut cb_args: Vec<Value> = Vec::with_capacity(caps.len() + 2);
                    for i in 0..caps.len() {
                        cb_args.push(match caps.get(i) {
                            Some(c) => s_val(c.as_str()),
                            None => Value::Undefined,
                        });
                    }
                    cb_args.push(Value::I32(m.start() as i32));
                    cb_args.push(s_val(&input));
                    let ret = ctx.invoke(&replacement_arg, &cb_args);
                    match ret {
                        Value::String(s) => out.push_str(s.as_ref()),
                        other => out.push_str(&format!("{}", other)),
                    }
                    last_end = m.end();
                    if !global { break; }
                }
                out.push_str(&input[last_end..]);
                return s_val(&out);
            }
            // String-form replacement: regex crate's $1/$2 syntax.
            let replacement = s_arg(args, 2);
            let result = if global {
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

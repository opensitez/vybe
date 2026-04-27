//! `node:util` — Node.js utility module.
//!
//! Reference: <https://nodejs.org/api/util.html>.
//!
//! Common across server-side JS runtimes (Node, Deno, Bun) — same tier
//! as `node:fs` / `node:path` / `node:os`. Not in browsers, not in
//! WinterCG Common Minimum API.
//!
//! Coverage:
//!   - `format(fmt, ...args)` / `formatWithOptions` — printf-style formatter
//!   - `inspect(obj, opts?)` — debug stringification
//!   - `isDeepStrictEqual(a, b)` — recursive equality
//!   - `stripVTControlCharacters(s)` — strip ANSI escape codes
//!   - `toUSVString(s)` — replace lone surrogates with U+FFFD
//!   - `parseArgs({ args, options, ... })` — CLI arg parser (Node 18+)
//!   - `types.is*(v)` — type predicates (`isArray`, `isMap`, `isSet`,
//!     `isRegExp`, `isDate`, `isPromise`, `isArrayBuffer`, etc.)
//!   - Legacy `isArray`/`isString`/etc. — deprecated but still exported
//!
//! Deferred (need callback/promise infrastructure):
//!   - `promisify(fn)`, `callbackify(fn)`
//!   - `inherits(ctor, super)` — legacy ES5 prototype linking
//!   - `deprecate(fn, msg)`, `debuglog(section)`

use std::sync::{Arc, Mutex};
use vybe_bytecode::value::{Object, ObjectKind, Value};
use vybe_bytecode::{HostContext, VM};

const MODULE: &str = "node:util";

pub fn register(vm: &mut VM) {
    register_format(vm);
    register_inspect(vm);
    register_deep_equal(vm);
    register_strip_vt(vm);
    register_to_usv(vm);
    register_parse_args(vm);
    register_types(vm);
    register_legacy_predicates(vm);
}

// ── format / formatWithOptions ───────────────────────────────────────

fn register_format(vm: &mut VM) {
    vm.register_host_fn(MODULE, "format", Box::new(format_impl));
    // `formatWithOptions(inspectOptions, fmt, ...args)` — same as
    // format but takes inspect options for `%o`/`%O`. MVP ignores
    // options and shifts args left by one before formatting.
    vm.register_host_fn(MODULE, "formatWithOptions", Box::new(|ctx, args| {
        if args.is_empty() { return Value::String(Arc::from("")); }
        format_impl(ctx, &args[1..])
    }));
}

/// Parse a Node `util.format` format string and substitute args.
///
/// Specifiers (per Node docs):
///   `%s` — String coercion (calls util.inspect for non-strings without ident)
///   `%d` — Number — int or float
///   `%i` — parseInt(value, 10) — truncates floats
///   `%f` — parseFloat(value)
///   `%j` — JSON.stringify
///   `%o` — util.inspect with `{ showHidden: true, depth: 4, showProxy: true }`
///   `%O` — util.inspect with default options
///   `%c` — CSS substitution (browser only — ignored, consumes one arg)
///   `%%` — literal `%`
///
/// Args without matching placeholders get appended space-separated.
fn format_impl(_ctx: &mut HostContext, args: &[Value]) -> Value {
    let fmt = match args.first() {
        Some(Value::String(s)) => s.to_string(),
        Some(other) => return Value::String(Arc::from(format!("{}", other).as_str())),
        None => return Value::String(Arc::from("")),
    };
    let extra = &args[1..];
    let mut out = String::with_capacity(fmt.len());
    let mut consumed = 0usize;
    let bytes = fmt.as_bytes();
    let mut i = 0;
    while i < bytes.len() {
        if bytes[i] == b'%' && i + 1 < bytes.len() {
            let spec = bytes[i + 1] as char;
            match spec {
                '%' => { out.push('%'); i += 2; continue; }
                's' | 'd' | 'i' | 'f' | 'j' | 'o' | 'O' | 'c' => {
                    if let Some(arg) = extra.get(consumed) {
                        match spec {
                            's' => out.push_str(&format_s(arg)),
                            'd' => out.push_str(&format_d(arg)),
                            'i' => out.push_str(&format_i(arg)),
                            'f' => out.push_str(&format_f(arg)),
                            'j' => out.push_str(&format_j(arg)),
                            'o' | 'O' => out.push_str(&inspect_value(arg, false)),
                            'c' => {} // CSS — consumed but no output
                            _ => unreachable!(),
                        }
                        consumed += 1;
                        i += 2;
                        continue;
                    }
                }
                _ => {}
            }
        }
        out.push(bytes[i] as char);
        i += 1;
    }
    // Spec: extra args beyond placeholders get appended space-separated.
    for arg in &extra[consumed..] {
        out.push(' ');
        out.push_str(&format_s(arg));
    }
    Value::String(Arc::from(out.as_str()))
}

fn format_s(v: &Value) -> String {
    match v {
        Value::String(s) => s.to_string(),
        Value::Object(_) => inspect_value(v, false),
        _ => format!("{}", v),
    }
}

fn format_d(v: &Value) -> String {
    match v {
        Value::I32(n) => n.to_string(),
        Value::I64(n) => n.to_string(),
        Value::F64(n) => {
            if n.is_finite() && *n == n.trunc() {
                format!("{}", *n as i64)
            } else {
                format!("{}", n)
            }
        }
        _ => "NaN".to_string(),
    }
}

fn format_i(v: &Value) -> String {
    match v {
        Value::I32(n) => n.to_string(),
        Value::I64(n) => n.to_string(),
        Value::F64(n) => format!("{}", n.trunc() as i64),
        Value::String(s) => s.trim().parse::<i64>().map(|n| n.to_string()).unwrap_or_else(|_| "NaN".into()),
        _ => "NaN".to_string(),
    }
}

fn format_f(v: &Value) -> String {
    match v {
        Value::F64(n) => format!("{}", n),
        Value::I32(n) => format!("{}", *n as f64),
        Value::I64(n) => format!("{}", *n as f64),
        Value::String(s) => s.trim().parse::<f64>().map(|n| format!("{}", n)).unwrap_or_else(|_| "NaN".into()),
        _ => "NaN".to_string(),
    }
}

fn format_j(v: &Value) -> String {
    json_stringify(v)
}

// ── inspect ──────────────────────────────────────────────────────────

fn register_inspect(vm: &mut VM) {
    vm.register_host_fn(MODULE, "inspect", Box::new(|_ctx, args| {
        let v = args.first().cloned().unwrap_or(Value::Undefined);
        Value::String(Arc::from(inspect_value(&v, false).as_str()))
    }));
}

/// Render a Value as a Node `util.inspect`-style debug string.
///
/// `quote_strings`: when true, wrap top-level strings in single quotes
/// (matches inspect behaviour on string args). When false, the string
/// renders raw — used for inner array/object elements where quoting
/// happens at the leaf.
fn inspect_value(v: &Value, _color: bool) -> String {
    inspect_inner(v, true, 0, 6)
}

fn inspect_inner(v: &Value, quote_strings: bool, depth: usize, max_depth: usize) -> String {
    match v {
        Value::Null => "null".into(),
        Value::Undefined => "undefined".into(),
        Value::Bool(b) => b.to_string(),
        Value::I32(n) => n.to_string(),
        Value::I64(n) => n.to_string(),
        Value::F64(n) => {
            if n.is_finite() && *n == n.trunc() {
                format!("{}", *n as i64)
            } else {
                format!("{}", n)
            }
        }
        Value::String(s) => {
            if quote_strings {
                format!("'{}'", s.replace('\\', "\\\\").replace('\'', "\\'"))
            } else {
                s.to_string()
            }
        }
        Value::Object(obj) => {
            if depth >= max_depth { return "[Object]".into(); }
            let o = obj.lock().unwrap();
            match &o.kind {
                ObjectKind::Array(elems) => {
                    if elems.is_empty() {
                        return "[]".into();
                    }
                    let parts: Vec<String> = elems.iter()
                        .map(|e| inspect_inner(e, true, depth + 1, max_depth))
                        .collect();
                    format!("[ {} ]", parts.join(", "))
                }
                ObjectKind::Map(m) => {
                    if m.is_empty() {
                        return "Map(0) {}".into();
                    }
                    let parts: Vec<String> = m.iter()
                        .map(|(k, v)| format!("{} => {}",
                            inspect_inner(k, true, depth + 1, max_depth),
                            inspect_inner(v, true, depth + 1, max_depth)))
                        .collect();
                    format!("Map({}) {{ {} }}", m.len(), parts.join(", "))
                }
                ObjectKind::Set(s) => {
                    if s.is_empty() {
                        return "Set(0) {}".into();
                    }
                    let parts: Vec<String> = s.iter()
                        .map(|v| inspect_inner(v, true, depth + 1, max_depth))
                        .collect();
                    format!("Set({}) {{ {} }}", s.len(), parts.join(", "))
                }
                _ => {
                    let visible: Vec<(&String, &Value)> = o.properties.iter()
                        .filter(|(k, _)| !k.starts_with("__"))
                        .collect();
                    if visible.is_empty() {
                        return "{}".into();
                    }
                    let parts: Vec<String> = visible.iter()
                        .map(|(k, v)| format!("{}: {}", k, inspect_inner(v, true, depth + 1, max_depth)))
                        .collect();
                    format!("{{ {} }}", parts.join(", "))
                }
            }
        }
        _ => format!("{}", v),
    }
}

// ── isDeepStrictEqual ────────────────────────────────────────────────

fn register_deep_equal(vm: &mut VM) {
    vm.register_host_fn(MODULE, "isDeepStrictEqual", Box::new(|_ctx, args| {
        let a = args.first().cloned().unwrap_or(Value::Undefined);
        let b = args.get(1).cloned().unwrap_or(Value::Undefined);
        Value::Bool(deep_strict_equal(&a, &b))
    }));
}

fn deep_strict_equal(a: &Value, b: &Value) -> bool {
    match (a, b) {
        (Value::Null, Value::Null) => true,
        (Value::Undefined, Value::Undefined) => true,
        (Value::Bool(x), Value::Bool(y)) => x == y,
        (Value::I32(x), Value::I32(y)) => x == y,
        (Value::I64(x), Value::I64(y)) => x == y,
        (Value::F64(x), Value::F64(y)) => x == y || (x.is_nan() && y.is_nan()),
        (Value::String(x), Value::String(y)) => x == y,
        (Value::Object(x), Value::Object(y)) => {
            if Arc::ptr_eq(x, y) { return true; }
            let xo = x.lock().unwrap();
            let yo = y.lock().unwrap();
            match (&xo.kind, &yo.kind) {
                (ObjectKind::Array(xa), ObjectKind::Array(ya)) => {
                    xa.len() == ya.len() && xa.iter().zip(ya.iter()).all(|(a, b)| deep_strict_equal(a, b))
                }
                (ObjectKind::Map(xm), ObjectKind::Map(ym)) => {
                    if xm.len() != ym.len() { return false; }
                    xm.iter().all(|(k, v)| ym.get(k).map_or(false, |yv| deep_strict_equal(v, yv)))
                }
                (ObjectKind::Set(xs), ObjectKind::Set(ys)) => {
                    xs.len() == ys.len() && xs.iter().all(|v| ys.contains(v))
                }
                _ => {
                    if xo.properties.len() != yo.properties.len() { return false; }
                    xo.properties.iter().all(|(k, v)| {
                        yo.properties.get(k).map_or(false, |yv| deep_strict_equal(v, yv))
                    })
                }
            }
        }
        _ => false,
    }
}

// ── stripVTControlCharacters ─────────────────────────────────────────

fn register_strip_vt(vm: &mut VM) {
    vm.register_host_fn(MODULE, "stripVTControlCharacters", Box::new(|_ctx, args| {
        let s = match args.first() {
            Some(Value::String(s)) => s.to_string(),
            Some(other) => format!("{}", other),
            None => return Value::String(Arc::from("")),
        };
        let mut out = String::with_capacity(s.len());
        let chars: Vec<char> = s.chars().collect();
        let mut i = 0;
        while i < chars.len() {
            // ANSI escape: ESC ([) ... letter (the final byte is a letter).
            if chars[i] == '\x1b' && i + 1 < chars.len() && chars[i + 1] == '[' {
                let mut j = i + 2;
                // Skip CSI parameter + intermediate bytes (digits, `;`, `?`, etc.).
                while j < chars.len() && !chars[j].is_ascii_alphabetic() {
                    j += 1;
                }
                // Skip the final byte (the letter terminator).
                if j < chars.len() {
                    j += 1;
                }
                i = j;
                continue;
            }
            out.push(chars[i]);
            i += 1;
        }
        Value::String(Arc::from(out.as_str()))
    }));
}

// ── toUSVString ──────────────────────────────────────────────────────
//
// Rust `String` is guaranteed UTF-8 — it cannot contain lone surrogates.
// So this is identity for any input we can hold. We only need to handle
// the case where the input is malformed bytes (which we can't represent
// as `Value::String` anyway) — passthrough is correct.

fn register_to_usv(vm: &mut VM) {
    vm.register_host_fn(MODULE, "toUSVString", Box::new(|_ctx, args| {
        let s = match args.first() {
            Some(Value::String(s)) => s.clone(),
            Some(other) => Arc::from(format!("{}", other).as_str()),
            None => Arc::from(""),
        };
        Value::String(s)
    }));
}

// ── parseArgs (Node 18+) ─────────────────────────────────────────────
//
// Spec: <https://nodejs.org/api/util.html#utilparseargsconfig>.
// Config:
//   - args: string[] — defaults to process.argv.slice(2)
//   - options: { name: { type: "string"|"boolean", short?, multiple?, default? } }
//   - allowPositionals: bool — default false
//   - strict: bool — default true
//   - tokens: bool — return raw token list (not implemented yet)
//
// Returns { values, positionals, tokens? }.

fn register_parse_args(vm: &mut VM) {
    vm.register_host_fn(MODULE, "parseArgs", Box::new(|_ctx, args| {
        let cfg = match args.first() {
            Some(Value::Object(c)) => c.clone(),
            _ => return new_parse_args_result(Vec::new(), Vec::new()),
        };

        let cfg_lock = cfg.lock().unwrap();
        let cli_args: Vec<String> = match cfg_lock.properties.get("args") {
            Some(Value::Object(arr)) => {
                let lo = arr.lock().unwrap();
                if let ObjectKind::Array(elems) = &lo.kind {
                    elems.iter().map(|v| match v {
                        Value::String(s) => s.to_string(),
                        other => format!("{}", other),
                    }).collect()
                } else { Vec::new() }
            }
            _ => Vec::new(),
        };

        // Collect option specs: name → (type, short?, multiple, default)
        let mut option_specs: indexmap::IndexMap<String, OptionSpec> = indexmap::IndexMap::new();
        if let Some(Value::Object(opts)) = cfg_lock.properties.get("options") {
            let opts_lock = opts.lock().unwrap();
            for (name, spec) in opts_lock.properties.iter() {
                if name.starts_with("__") { continue; }
                let mut s = OptionSpec::default();
                if let Value::Object(o) = spec {
                    let ol = o.lock().unwrap();
                    if let Some(Value::String(t)) = ol.properties.get("type") {
                        s.is_boolean = t.as_ref() == "boolean";
                    }
                    if let Some(Value::String(short)) = ol.properties.get("short") {
                        s.short = Some(short.to_string());
                    }
                    if let Some(Value::Bool(m)) = ol.properties.get("multiple") {
                        s.multiple = *m;
                    }
                }
                option_specs.insert(name.clone(), s);
            }
        }

        let allow_positionals = matches!(
            cfg_lock.properties.get("allowPositionals"),
            Some(Value::Bool(true))
        );
        drop(cfg_lock);

        // Walk args
        let mut values: indexmap::IndexMap<String, Value> = indexmap::IndexMap::new();
        let mut positionals: Vec<Value> = Vec::new();
        let mut i = 0;
        while i < cli_args.len() {
            let arg = &cli_args[i];
            if let Some(rest) = arg.strip_prefix("--") {
                // Long option: --name=value or --name value
                let (name, inline_value) = match rest.find('=') {
                    Some(eq) => (rest[..eq].to_string(), Some(rest[eq + 1..].to_string())),
                    None => (rest.to_string(), None),
                };
                if let Some(spec) = option_specs.get(&name) {
                    if spec.is_boolean {
                        values.insert(name, Value::Bool(true));
                        i += 1;
                    } else {
                        let v = inline_value.or_else(|| {
                            let next = cli_args.get(i + 1).cloned();
                            if next.is_some() { i += 1; }
                            next
                        }).unwrap_or_default();
                        values.insert(name, Value::String(Arc::from(v.as_str())));
                        i += 1;
                    }
                } else {
                    i += 1;
                }
            } else if let Some(short) = arg.strip_prefix('-').filter(|s| !s.is_empty() && !s.starts_with('-')) {
                // Short option (single-char form)
                let resolved = option_specs.iter().find_map(|(name, spec)| {
                    if spec.short.as_deref() == Some(short) { Some((name.clone(), spec.clone())) } else { None }
                });
                if let Some((name, spec)) = resolved {
                    if spec.is_boolean {
                        values.insert(name, Value::Bool(true));
                    } else if let Some(next) = cli_args.get(i + 1).cloned() {
                        values.insert(name, Value::String(Arc::from(next.as_str())));
                        i += 1;
                    }
                }
                i += 1;
            } else if allow_positionals {
                positionals.push(Value::String(Arc::from(arg.as_str())));
                i += 1;
            } else {
                i += 1;
            }
        }

        let values_vec: Vec<(String, Value)> = values.into_iter().collect();
        new_parse_args_result(values_vec, positionals)
    }));
}

#[derive(Default, Clone)]
struct OptionSpec {
    is_boolean: bool,
    short: Option<String>,
    multiple: bool,
}

fn new_parse_args_result(values: Vec<(String, Value)>, positionals: Vec<Value>) -> Value {
    let mut values_obj = Object::new();
    for (k, v) in values {
        values_obj.properties.insert(k, v);
    }
    let positionals_arr = Object::new_array(positionals);
    let mut result = Object::new();
    result.properties.insert("values".into(), Value::Object(Arc::new(Mutex::new(values_obj))));
    result.properties.insert("positionals".into(), Value::Object(Arc::new(Mutex::new(positionals_arr))));
    Value::Object(Arc::new(Mutex::new(result)))
}

// ── types.* — type predicates ────────────────────────────────────────
//
// Node namespaces these under `util.types`. Vybe registers each under
// the dotted name (`types.isArray`, `types.isMap`, ...) since the host
// registry uses a flat (module, name) key.

fn register_types(vm: &mut VM) {
    // Array / Map / Set
    vm.register_host_fn(MODULE, "types.isArray", Box::new(|_ctx, args| {
        Value::Bool(matches!(args.first(), Some(Value::Object(o))
            if matches!(o.lock().unwrap().kind, ObjectKind::Array(_))))
    }));
    vm.register_host_fn(MODULE, "types.isMap", Box::new(|_ctx, args| {
        Value::Bool(matches!(args.first(), Some(Value::Object(o))
            if matches!(o.lock().unwrap().kind, ObjectKind::Map(_))))
    }));
    vm.register_host_fn(MODULE, "types.isSet", Box::new(|_ctx, args| {
        Value::Bool(matches!(args.first(), Some(Value::Object(o))
            if matches!(o.lock().unwrap().kind, ObjectKind::Set(_))))
    }));

    // Buffers / DataView / TypedArrays — all use ObjectKind variants
    vm.register_host_fn(MODULE, "types.isArrayBuffer", Box::new(|_ctx, args| {
        Value::Bool(matches!(args.first(), Some(Value::Object(o))
            if matches!(o.lock().unwrap().kind, ObjectKind::ArrayBuffer(_))))
    }));
    vm.register_host_fn(MODULE, "types.isSharedArrayBuffer", Box::new(|_ctx, args| {
        // No separate ObjectKind for SharedArrayBuffer yet; ArrayBuffer
        // covers both. Predicate returns true for any ArrayBuffer.
        Value::Bool(matches!(args.first(), Some(Value::Object(o))
            if matches!(o.lock().unwrap().kind, ObjectKind::ArrayBuffer(_))))
    }));
    vm.register_host_fn(MODULE, "types.isAnyArrayBuffer", Box::new(|_ctx, args| {
        Value::Bool(matches!(args.first(), Some(Value::Object(o))
            if matches!(o.lock().unwrap().kind, ObjectKind::ArrayBuffer(_))))
    }));
    vm.register_host_fn(MODULE, "types.isDataView", Box::new(|_ctx, args| {
        Value::Bool(is_typed_kind(args, "DataView"))
    }));
    vm.register_host_fn(MODULE, "types.isTypedArray", Box::new(|_ctx, args| {
        Value::Bool(matches!(args.first(), Some(Value::Object(o))
            if matches!(o.lock().unwrap().kind, ObjectKind::TypedArray(_))))
    }));

    // Specific typed arrays — distinguished by `__type` stamp
    for kind in &[
        "Int8Array", "Uint8Array", "Uint8ClampedArray",
        "Int16Array", "Uint16Array",
        "Int32Array", "Uint32Array",
        "Float32Array", "Float64Array",
        "BigInt64Array", "BigUint64Array",
    ] {
        let kind_name = kind.to_string();
        let predicate_name = format!("types.is{}", kind);
        vm.register_host_fn(MODULE, &predicate_name, Box::new(move |_ctx, args| {
            Value::Bool(is_typed_kind(args, &kind_name))
        }));
    }

    // Date / RegExp / Promise / Error — recognized via __type stamp
    for (predicate, type_tag) in &[
        ("types.isDate", "Date"),
        ("types.isRegExp", "RegExp"),
        ("types.isPromise", "Promise"),
        ("types.isNativeError", "Error"),
    ] {
        let tag = type_tag.to_string();
        let pred = predicate.to_string();
        vm.register_host_fn(MODULE, &pred, Box::new(move |_ctx, args| {
            Value::Bool(is_typed_kind(args, &tag))
        }));
    }

    // Function-shape predicates. Vybe's Function value-shape doesn't
    // carry the async/generator flag (that lives on the underlying
    // Chunk). Without VM access from a host fn we can't reach the
    // chunk, so for now these predicates check for a `__type` stamp
    // ("AsyncFunction" / "GeneratorFunction") that the JS class
    // normalizer can install on async/generator function objects.
    vm.register_host_fn(MODULE, "types.isAsyncFunction", Box::new(|_ctx, args| {
        Value::Bool(is_typed_kind(args, "AsyncFunction"))
    }));
    vm.register_host_fn(MODULE, "types.isGeneratorFunction", Box::new(|_ctx, args| {
        Value::Bool(is_typed_kind(args, "GeneratorFunction"))
    }));
    vm.register_host_fn(MODULE, "types.isGeneratorObject", Box::new(|_ctx, args| {
        Value::Bool(is_typed_kind(args, "Generator"))
    }));

    // Iterator predicates — Vybe represents iterators as ordinary objects
    // with __type stamps ("MapIterator", "SetIterator").
    vm.register_host_fn(MODULE, "types.isMapIterator", Box::new(|_ctx, args| {
        Value::Bool(is_typed_kind(args, "MapIterator"))
    }));
    vm.register_host_fn(MODULE, "types.isSetIterator", Box::new(|_ctx, args| {
        Value::Bool(is_typed_kind(args, "SetIterator"))
    }));

    // WeakMap / WeakSet
    vm.register_host_fn(MODULE, "types.isWeakMap", Box::new(|_ctx, args| {
        Value::Bool(is_typed_kind(args, "WeakMap"))
    }));
    vm.register_host_fn(MODULE, "types.isWeakSet", Box::new(|_ctx, args| {
        Value::Bool(is_typed_kind(args, "WeakSet"))
    }));

    // Boxed primitives — Vybe doesn't box primitives by default; these
    // return false unless the value is explicitly boxed via `Object(x)`
    // (which we don't do today). Predicates exist for spec faithfulness.
    for tag in &["Boolean", "Number", "String", "Symbol", "BigInt"] {
        let stamp = tag.to_string();
        let pred = format!("types.is{}Object", tag);
        vm.register_host_fn(MODULE, &pred, Box::new(move |_ctx, args| {
            Value::Bool(is_typed_kind(args, &stamp))
        }));
    }

    // BoxedPrimitive — true if any of the boxed predicates would be true.
    vm.register_host_fn(MODULE, "types.isBoxedPrimitive", Box::new(|_ctx, args| {
        Value::Bool(
            is_typed_kind(args, "Boolean") ||
            is_typed_kind(args, "Number") ||
            is_typed_kind(args, "String") ||
            is_typed_kind(args, "Symbol") ||
            is_typed_kind(args, "BigInt")
        )
    }));

    // Misc — currently no support for these in Vybe's type system, so
    // they always return false. Listed for spec completeness.
    for pred in &[
        "types.isArgumentsObject",
        "types.isExternal",
        "types.isProxy",
        "types.isModuleNamespaceObject",
        "types.isCryptoKey",
        "types.isKeyObject",
    ] {
        let p = pred.to_string();
        vm.register_host_fn(MODULE, &p, Box::new(|_ctx, _args| Value::Bool(false)));
    }
}

/// Helper: true iff arg[0] is an Object whose `__type` property
/// (case-sensitive) equals `tag`. Used by every `types.is*` predicate
/// that recognizes a stamped type.
fn is_typed_kind(args: &[Value], tag: &str) -> bool {
    if let Some(Value::Object(o)) = args.first() {
        let lock = o.lock().unwrap();
        if let Some(Value::String(t)) = lock.properties.get("__type") {
            return t.as_ref() == tag;
        }
    }
    false
}

// ── Legacy is* predicates (deprecated but still shipped) ─────────────
//
// Per Node docs these are "Deprecated" but functional — many real-world
// codebases still use them, so Node maintains them for back-compat. We
// do the same.

fn register_legacy_predicates(vm: &mut VM) {
    vm.register_host_fn(MODULE, "isArray", Box::new(|_ctx, args| {
        Value::Bool(matches!(args.first(), Some(Value::Object(o))
            if matches!(o.lock().unwrap().kind, ObjectKind::Array(_))))
    }));
    vm.register_host_fn(MODULE, "isString", Box::new(|_ctx, args| {
        Value::Bool(matches!(args.first(), Some(Value::String(_))))
    }));
    vm.register_host_fn(MODULE, "isNumber", Box::new(|_ctx, args| {
        Value::Bool(matches!(args.first(), Some(Value::I32(_) | Value::I64(_) | Value::F64(_))))
    }));
    vm.register_host_fn(MODULE, "isBoolean", Box::new(|_ctx, args| {
        Value::Bool(matches!(args.first(), Some(Value::Bool(_))))
    }));
    vm.register_host_fn(MODULE, "isNull", Box::new(|_ctx, args| {
        Value::Bool(matches!(args.first(), Some(Value::Null)))
    }));
    vm.register_host_fn(MODULE, "isUndefined", Box::new(|_ctx, args| {
        Value::Bool(matches!(args.first(), Some(Value::Undefined)))
    }));
    vm.register_host_fn(MODULE, "isNullOrUndefined", Box::new(|_ctx, args| {
        Value::Bool(matches!(args.first(), Some(Value::Null | Value::Undefined)))
    }));
    vm.register_host_fn(MODULE, "isObject", Box::new(|_ctx, args| {
        // Per Node: `isObject(null)` returns false (matches `typeof null === "object"`
        // but the legacy util.isObject excludes null per docs).
        Value::Bool(matches!(args.first(), Some(Value::Object(_))))
    }));
    vm.register_host_fn(MODULE, "isPrimitive", Box::new(|_ctx, args| {
        Value::Bool(matches!(args.first(),
            Some(Value::Null | Value::Undefined | Value::Bool(_)
                 | Value::I32(_) | Value::I64(_) | Value::F64(_)
                 | Value::String(_) | Value::Symbol(_) | Value::BigInt(_))))
    }));
    vm.register_host_fn(MODULE, "isSymbol", Box::new(|_ctx, args| {
        Value::Bool(matches!(args.first(), Some(Value::Symbol(_))))
    }));
    vm.register_host_fn(MODULE, "isFunction", Box::new(|_ctx, args| {
        Value::Bool(matches!(args.first(), Some(Value::Object(o)) if {
            matches!(o.lock().unwrap().kind, ObjectKind::Function(_))
        }))
    }));
    vm.register_host_fn(MODULE, "isDate", Box::new(|_ctx, args| {
        Value::Bool(is_typed_kind(args, "Date"))
    }));
    vm.register_host_fn(MODULE, "isRegExp", Box::new(|_ctx, args| {
        Value::Bool(is_typed_kind(args, "RegExp"))
    }));
    vm.register_host_fn(MODULE, "isError", Box::new(|_ctx, args| {
        Value::Bool(is_typed_kind(args, "Error"))
    }));
    vm.register_host_fn(MODULE, "isBuffer", Box::new(|_ctx, args| {
        // Node Buffer is an ArrayBuffer view in spirit; we check the
        // ArrayBuffer kind since Vybe doesn't have a separate Buffer type.
        Value::Bool(matches!(args.first(), Some(Value::Object(o))
            if matches!(o.lock().unwrap().kind, ObjectKind::ArrayBuffer(_))))
    }));
}

// ── Local JSON stringify (for %j format spec) ────────────────────────
//
// Mirrors `ecma:json.stringify` but inlined here so format doesn't have
// to invoke through the host registry. Same output shape as that fn.

fn json_stringify(v: &Value) -> String {
    match v {
        Value::Null | Value::Undefined => "null".into(),
        Value::Bool(b) => b.to_string(),
        Value::I32(n) => n.to_string(),
        Value::I64(n) => n.to_string(),
        Value::F64(n) => {
            if n.is_nan() || n.is_infinite() { "null".into() }
            else if *n == (*n as i64) as f64 { format!("{}", *n as i64) }
            else { format!("{}", n) }
        }
        Value::String(s) => format!("\"{}\"", s.replace('\\', "\\\\").replace('"', "\\\"").replace('\n', "\\n")),
        Value::Object(obj) => {
            let o = obj.lock().unwrap();
            match &o.kind {
                ObjectKind::Array(elems) => {
                    let parts: Vec<String> = elems.iter().map(json_stringify).collect();
                    format!("[{}]", parts.join(","))
                }
                ObjectKind::Map(m) => {
                    let parts: Vec<String> = m.iter()
                        .map(|(k, v)| format!("\"{}\":{}",
                            match k { Value::String(s) => s.to_string(), o => format!("{}", o) },
                            json_stringify(v)))
                        .collect();
                    format!("{{{}}}", parts.join(","))
                }
                _ => {
                    let parts: Vec<String> = o.properties.iter()
                        .filter(|(k, _)| !k.starts_with("__"))
                        .map(|(k, v)| format!("\"{}\":{}", k, json_stringify(v)))
                        .collect();
                    format!("{{{}}}", parts.join(","))
                }
            }
        }
        _ => "null".into(),
    }
}

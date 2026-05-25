//! # `ecma:json` host handlers
//!
//! `JSON.stringify` / `JSON.parse` per ECMA-262 §25.5.
//!
//! Implementation notes:
//!   - `stringify` handles all value kinds in our VM: Array (JS
//!     array), Map (serializes as empty object per spec — Maps aren't
//!     natively stringifiable), Set (same), Object (property bag),
//!     TypedArray (numeric-indexed array of elements), ArrayBuffer
//!     (serialized as `{}` per spec), primitives with the usual JS
//!     rules.
//!   - `parse` produces Array / Object / Value primitives — never
//!     Map/Set/TypedArray (the spec says JSON.parse always materializes
//!     JS Objects and Arrays).
//!   - NaN / Infinity stringify to `"null"` per spec.
//!   - `undefined` elements in Arrays stringify as `"null"`;
//!     `undefined` properties in Objects are **omitted** (spec).
//!   - Circular references: we detect via a visited-set and throw
//!     (MVP: returns an error string instead of trapping).
//!   - The `replacer` and `space` arguments are currently ignored —
//!     Phase B5 follow-up.
//!
//! See `JS_BUILTIN_CONVENTIONS.md`.

use std::collections::HashSet;
use std::sync::{Arc, Mutex};
use vybe_bytecode::value::{Object, ObjectKind, TypedElemKind, Value};
use vybe_bytecode::VM;

pub fn register(vm: &mut VM) {
    vm.register_host_fn("ecma:json", "stringify",
        Box::new(|_ctx, args| {
            let value = args.first().cloned().unwrap_or(Value::Undefined);
            // Spec: stringify(undefined) returns undefined, not "undefined"
            if matches!(value, Value::Undefined) {
                return Value::Undefined;
            }
            // Spec: stringify(Symbol) returns undefined
            if matches!(value, Value::Symbol(_)) {
                return Value::Undefined;
            }
            let mut visited: HashSet<usize> = HashSet::new();
            let s = stringify(&value, &mut visited);
            Value::String(Arc::from(s.as_str()))
        }));

    vm.register_host_fn("ecma:json", "parse",
        Box::new(|_ctx, args| {
            let text: String = match args.first() {
                Some(Value::String(s)) => s.to_string(),
                Some(other) => format!("{}", other),
                None => return Value::Undefined,
            };
            parse_json(&text).unwrap_or(Value::Null)
        }));
}

// ── Stringify ──────────────────────────────────────────────────────────

fn stringify(v: &Value, visited: &mut HashSet<usize>) -> String {
    match v {
        Value::Null | Value::Undefined => "null".to_string(),
        Value::Bool(b) => if *b { "true".into() } else { "false".into() },
        Value::I32(n) => n.to_string(),
        Value::I64(n) => n.to_string(),
        Value::F64(n) => {
            if n.is_nan() || n.is_infinite() {
                "null".to_string()  // ECMA-262 §25.5.2 — NaN/Infinity stringify as null
            } else if *n == (*n as i64) as f64 && n.abs() < 1e16 {
                (*n as i64).to_string()
            } else {
                n.to_string()
            }
        }
        Value::String(s) => quote_string(s),
        Value::BigInt(_) => "null".into(), // spec says TypeError; MVP emits null
        Value::Symbol(_) => "null".into(),
        Value::V128(_) => "null".into(),
        Value::WeakRef(_) => "null".into(),
        Value::Object(obj) => stringify_object(obj, visited),
    }
}

fn stringify_object(obj: &Arc<Mutex<Object>>, visited: &mut HashSet<usize>) -> String {
    let id = Arc::as_ptr(obj) as usize;
    if !visited.insert(id) {
        // Cycle detected. Per spec we should throw TypeError; MVP
        // emits a sentinel string.
        return "null".to_string();
    }
    let result = {
        let o = obj.lock().unwrap();
        match &o.kind {
            ObjectKind::Array(elems) => stringify_array(elems, visited),
            ObjectKind::TypedArray(ta) => stringify_typed_array(ta, visited),
            ObjectKind::Map(_) | ObjectKind::Set(_) => {
                // Spec: Map/Set serialize as {} (no enumerable own
                // properties). Matches v8.
                "{}".to_string()
            }
            ObjectKind::ArrayBuffer(_) => "{}".to_string(),
            ObjectKind::Function(_) | ObjectKind::HostFunction(_) => {
                // Spec: functions are omitted (handled by caller when
                // nested in Object/Array); top-level returns undefined.
                "null".to_string()
            }
            ObjectKind::Continuation(_) => {
                // Continuations don't serialize — match the function
                // treatment above (no enumerable own data).
                "null".to_string()
            }
            ObjectKind::Ordinary => {
                // Date receives toJSON treatment per ECMA-262 §25.5.1.1:
                // if `__type=Date`, serialize as the ISO string. Same logic
                // V8 uses to make `JSON.stringify(d)` return `"2026-..."`.
                // Read __time directly while we hold the lock — calling
                // `dispatch_date_method` would re-lock and deadlock.
                let is_date = matches!(
                    o.properties.get("__type"),
                    Some(Value::String(s)) if s.as_ref() == "Date"
                );
                if is_date {
                    let ms = o.properties.get("__time")
                        .map(|v| v.as_f64())
                        .unwrap_or(f64::NAN);
                    let iso = crate::ecma::date::format_iso_from_ms(ms);
                    return match iso {
                        Some(s) => format!("\"{}\"", s),
                        None => "null".to_string(),
                    };
                }
                stringify_ordinary(&o, visited)
            }
            // Module Namespace Objects serialize their exports like an
            // Ordinary object — functions (the common case) get dropped
            // per the same "value is a function → undefined" rule.
            ObjectKind::ModuleNamespace => {
                stringify_ordinary(&o, visited)
            }
        }
    };
    visited.remove(&id);
    result
}

fn stringify_array(elems: &[Value], visited: &mut HashSet<usize>) -> String {
    let parts: Vec<String> = elems.iter().map(|v| {
        match v {
            // Functions / symbols / undefined in arrays serialize as "null"
            Value::Undefined | Value::Symbol(_) => "null".to_string(),
            Value::Object(o) => {
                let is_fn = {
                    let lock = o.lock().unwrap();
                    matches!(lock.kind, ObjectKind::Function(_) | ObjectKind::HostFunction(_))
                };
                if is_fn { "null".into() } else { stringify(v, visited) }
            }
            _ => stringify(v, visited),
        }
    }).collect();
    format!("[{}]", parts.join(","))
}

fn stringify_typed_array(ta: &vybe_bytecode::value::TypedArrayState,
                          _visited: &mut HashSet<usize>) -> String {
    // Typed arrays stringify as the comma-joined element values
    // wrapped in an object shape — actually JSON.stringify on a typed
    // array produces a plain object with numeric-string keys. v8:
    //   JSON.stringify(new Int32Array([1,2,3])) === '{"0":1,"1":2,"2":3}'
    let buf = ta.buffer.lock().unwrap();
    let bpe = ta.elem.bytes_per_element();
    let available_elems = if ta.byte_offset >= buf.len() { 0 }
        else { (buf.len() - ta.byte_offset) / bpe };
    let length = ta.length.min(available_elems);
    let mut out = String::from("{");
    for i in 0..length {
        if i > 0 { out.push(','); }
        out.push_str(&format!("\"{}\":", i));
        let abs = ta.byte_offset + i * bpe;
        let val_str = match ta.elem {
            TypedElemKind::I8  => (buf[abs] as i8).to_string(),
            TypedElemKind::U8 | TypedElemKind::U8Clamped => buf[abs].to_string(),
            TypedElemKind::I16 => {
                let b = [buf[abs], buf[abs + 1]];
                i16::from_le_bytes(b).to_string()
            }
            TypedElemKind::U16 => {
                let b = [buf[abs], buf[abs + 1]];
                u16::from_le_bytes(b).to_string()
            }
            TypedElemKind::I32 => {
                let mut b = [0u8; 4]; b.copy_from_slice(&buf[abs..abs + 4]);
                i32::from_le_bytes(b).to_string()
            }
            TypedElemKind::U32 => {
                let mut b = [0u8; 4]; b.copy_from_slice(&buf[abs..abs + 4]);
                u32::from_le_bytes(b).to_string()
            }
            TypedElemKind::F32 => {
                let mut b = [0u8; 4]; b.copy_from_slice(&buf[abs..abs + 4]);
                let f = f32::from_le_bytes(b);
                if f.is_nan() || f.is_infinite() { "null".into() } else { f.to_string() }
            }
            TypedElemKind::F64 => {
                let mut b = [0u8; 8]; b.copy_from_slice(&buf[abs..abs + 8]);
                let f = f64::from_le_bytes(b);
                if f.is_nan() || f.is_infinite() { "null".into() } else { f.to_string() }
            }
            TypedElemKind::BigI64 => {
                let mut b = [0u8; 8]; b.copy_from_slice(&buf[abs..abs + 8]);
                i64::from_le_bytes(b).to_string()
            }
            TypedElemKind::BigU64 => {
                let mut b = [0u8; 8]; b.copy_from_slice(&buf[abs..abs + 8]);
                u64::from_le_bytes(b).to_string()
            }
        };
        out.push_str(&val_str);
    }
    out.push('}');
    out
}

fn ordinary_ordered_keys(o: &Object) -> Vec<String> {
    let tracked: Option<Vec<String>> = o.properties.get("__keys").and_then(|value| {
        let Value::Object(arr) = value else {
            return None;
        };
        let guard = arr.lock().unwrap();
        let ObjectKind::Array(ref elems) = guard.kind else {
            return None;
        };
        Some(
            elems.iter()
                .filter_map(|elem| match elem {
                    Value::String(key) if o.properties.contains_key(key.as_ref()) => Some(key.to_string()),
                    _ => None,
                })
                .collect(),
        )
    });

    let live: Vec<String> = o.properties.keys().cloned().collect();
    match tracked {
        Some(mut keys) => {
            let mut seen: HashSet<String> = keys.iter().cloned().collect();
            for key in live {
                if seen.insert(key.clone()) {
                    keys.push(key);
                }
            }
            keys
        }
        None => live,
    }
}

fn stringify_ordinary(o: &Object, visited: &mut HashSet<usize>) -> String {
    let mut parts: Vec<String> = Vec::new();
    for k in ordinary_ordered_keys(o) {
        let Some(v) = o.properties.get(&k) else {
            continue;
        };
        // Skip internal __vybe_* bookkeeping properties.
        if k.starts_with("__") { continue; }
        // ECMA-262 §25.5.2.5: Symbol-keyed properties are not serialized.
        // Our VM stores them as "Symbol(<desc>)" string keys.
        if k.starts_with("Symbol(") && k.ends_with(')') { continue; }
        // Skip undefined / function values per spec (omitted, not
        // serialized as "null").
        match v {
            Value::Undefined | Value::Symbol(_) => continue,
            Value::Object(inner) => {
                let is_fn = {
                    let lock = inner.lock().unwrap();
                    matches!(lock.kind, ObjectKind::Function(_) | ObjectKind::HostFunction(_))
                };
                if is_fn { continue; }
            }
            _ => {}
        }
        parts.push(format!("{}:{}", quote_string(&k), stringify(v, visited)));
    }
    format!("{{{}}}", parts.join(","))
}

fn quote_string(s: &str) -> String {
    let mut out = String::with_capacity(s.len() + 2);
    out.push('"');
    for c in s.chars() {
        match c {
            '"' => out.push_str("\\\""),
            '\\' => out.push_str("\\\\"),
            '\n' => out.push_str("\\n"),
            '\r' => out.push_str("\\r"),
            '\t' => out.push_str("\\t"),
            '\x08' => out.push_str("\\b"),
            '\x0C' => out.push_str("\\f"),
            c if (c as u32) < 0x20 => {
                out.push_str(&format!("\\u{:04x}", c as u32));
            }
            c => out.push(c),
        }
    }
    out.push('"');
    out
}

// ── Parse ──────────────────────────────────────────────────────────────

struct Parser<'a> {
    src: &'a [u8],
    pos: usize,
}

fn parse_json(text: &str) -> Option<Value> {
    let mut p = Parser { src: text.as_bytes(), pos: 0 };
    p.skip_whitespace();
    let v = p.parse_value()?;
    p.skip_whitespace();
    if p.pos != p.src.len() {
        // Trailing content — per spec this is a SyntaxError; MVP returns null.
        return None;
    }
    Some(v)
}

impl<'a> Parser<'a> {
    fn skip_whitespace(&mut self) {
        while self.pos < self.src.len() {
            match self.src[self.pos] {
                b' ' | b'\t' | b'\n' | b'\r' => self.pos += 1,
                _ => break,
            }
        }
    }

    fn peek(&self) -> Option<u8> {
        self.src.get(self.pos).copied()
    }

    fn parse_value(&mut self) -> Option<Value> {
        self.skip_whitespace();
        match self.peek()? {
            b'{' => self.parse_object(),
            b'[' => self.parse_array(),
            b'"' => self.parse_string().map(|s| Value::String(Arc::from(s.as_str()))),
            b't' | b'f' => self.parse_bool(),
            b'n' => self.parse_null(),
            b'-' | b'0'..=b'9' => self.parse_number(),
            _ => None,
        }
    }

    fn parse_null(&mut self) -> Option<Value> {
        if self.src[self.pos..].starts_with(b"null") {
            self.pos += 4;
            Some(Value::Null)
        } else {
            None
        }
    }

    fn parse_bool(&mut self) -> Option<Value> {
        if self.src[self.pos..].starts_with(b"true") {
            self.pos += 4;
            Some(Value::Bool(true))
        } else if self.src[self.pos..].starts_with(b"false") {
            self.pos += 5;
            Some(Value::Bool(false))
        } else {
            None
        }
    }

    fn parse_number(&mut self) -> Option<Value> {
        let start = self.pos;
        if self.src.get(self.pos) == Some(&b'-') { self.pos += 1; }
        while self.pos < self.src.len() {
            let c = self.src[self.pos];
            if !(c.is_ascii_digit() || c == b'.' || c == b'e' || c == b'E'
                || c == b'+' || c == b'-')
            {
                break;
            }
            self.pos += 1;
        }
        let s = std::str::from_utf8(&self.src[start..self.pos]).ok()?;
        // Try i64 first for exact integer preservation, then f64.
        if !s.contains('.') && !s.contains('e') && !s.contains('E') {
            if let Ok(n) = s.parse::<i64>() {
                // Fit into i32 if possible — matches v8's tagging
                // preference for small integers.
                if n >= i32::MIN as i64 && n <= i32::MAX as i64 {
                    return Some(Value::I32(n as i32));
                }
                return Some(Value::I64(n));
            }
        }
        s.parse::<f64>().ok().map(Value::F64)
    }

    fn parse_string(&mut self) -> Option<String> {
        if self.peek()? != b'"' { return None; }
        self.pos += 1;
        let mut out = String::new();
        while self.pos < self.src.len() {
            let c = self.src[self.pos];
            match c {
                b'"' => {
                    self.pos += 1;
                    return Some(out);
                }
                b'\\' => {
                    self.pos += 1;
                    let esc = *self.src.get(self.pos)?;
                    self.pos += 1;
                    match esc {
                        b'"' => out.push('"'),
                        b'\\' => out.push('\\'),
                        b'/' => out.push('/'),
                        b'n' => out.push('\n'),
                        b't' => out.push('\t'),
                        b'r' => out.push('\r'),
                        b'b' => out.push('\x08'),
                        b'f' => out.push('\x0C'),
                        b'u' => {
                            if self.pos + 4 > self.src.len() { return None; }
                            let hex = std::str::from_utf8(&self.src[self.pos..self.pos + 4]).ok()?;
                            let code = u32::from_str_radix(hex, 16).ok()?;
                            self.pos += 4;
                            if let Some(ch) = char::from_u32(code) {
                                out.push(ch);
                            }
                        }
                        _ => return None,
                    }
                }
                _ => {
                    // Copy UTF-8 bytes up to the next special char.
                    // Cheapest: push one byte if ASCII, else advance
                    // and copy the full char via from_utf8.
                    let remaining = &self.src[self.pos..];
                    // Find the end of this char's byte sequence.
                    let char_len = utf8_char_len(remaining[0]);
                    if char_len == 0 || self.pos + char_len > self.src.len() {
                        return None;
                    }
                    let chunk = std::str::from_utf8(&remaining[..char_len]).ok()?;
                    out.push_str(chunk);
                    self.pos += char_len;
                }
            }
        }
        None // unterminated string
    }

    fn parse_array(&mut self) -> Option<Value> {
        if self.peek()? != b'[' { return None; }
        self.pos += 1;
        let mut elems: Vec<Value> = Vec::new();
        self.skip_whitespace();
        if self.peek() == Some(b']') {
            self.pos += 1;
            return Some(Value::Object(Arc::new(Mutex::new(Object::new_array(elems)))));
        }
        loop {
            elems.push(self.parse_value()?);
            self.skip_whitespace();
            match self.peek()? {
                b',' => { self.pos += 1; }
                b']' => { self.pos += 1; break; }
                _ => return None,
            }
        }
        Some(Value::Object(Arc::new(Mutex::new(Object::new_array(elems)))))
    }

    fn parse_object(&mut self) -> Option<Value> {
        if self.peek()? != b'{' { return None; }
        self.pos += 1;
        let mut obj = Object::new();
        self.skip_whitespace();
        if self.peek() == Some(b'}') {
            self.pos += 1;
            return Some(Value::Object(Arc::new(Mutex::new(obj))));
        }
        loop {
            self.skip_whitespace();
            let key = self.parse_string()?;
            self.skip_whitespace();
            if self.peek() != Some(b':') { return None; }
            self.pos += 1;
            let val = self.parse_value()?;
            obj.properties.insert(key, val);
            self.skip_whitespace();
            match self.peek()? {
                b',' => { self.pos += 1; }
                b'}' => { self.pos += 1; break; }
                _ => return None,
            }
        }
        Some(Value::Object(Arc::new(Mutex::new(obj))))
    }
}

/// UTF-8 leading-byte → sequence length. Returns 0 on invalid.
fn utf8_char_len(b: u8) -> usize {
    if b & 0x80 == 0 { 1 }
    else if b & 0xE0 == 0xC0 { 2 }
    else if b & 0xF0 == 0xE0 { 3 }
    else if b & 0xF8 == 0xF0 { 4 }
    else { 0 }
}

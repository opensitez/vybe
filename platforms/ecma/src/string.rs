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
//! Cross-language note: all language profiles (VB/C#/PHP/Python/etc.)
//! now target this surface directly; every language sees the same
//! JS-runtime semantics.

use std::borrow::Cow;
use std::fmt::Write as _;
use std::sync::{Arc, Mutex, OnceLock};
use unicode_normalization::UnicodeNormalization;
use vybe_runtime::value::{Object, ObjectKind};
use vybe_runtime::vm::HostFnDecl;
use vybe_runtime::{FuncSig, HostContext, Param, VM, ValType, Value};

static STRING_PROTOTYPE: OnceLock<Arc<Mutex<Object>>> = OnceLock::new();
const HEX_UPPER: &[u8; 16] = b"0123456789ABCDEF";

#[inline]
fn push_percent_byte(out: &mut String, byte: u8) {
    out.push('%');
    out.push(HEX_UPPER[(byte >> 4) as usize] as char);
    out.push(HEX_UPPER[(byte & 0x0f) as usize] as char);
}

#[inline]
fn push_percent_u16(out: &mut String, code: u32) {
    out.push('%');
    out.push('u');
    out.push(HEX_UPPER[((code >> 12) & 0x0f) as usize] as char);
    out.push(HEX_UPPER[((code >> 8) & 0x0f) as usize] as char);
    out.push(HEX_UPPER[((code >> 4) & 0x0f) as usize] as char);
    out.push(HEX_UPPER[(code & 0x0f) as usize] as char);
}

fn push_string_value(out: &mut String, value: &Value) {
    match value {
        Value::String(text) => out.push_str(text),
        other => {
            let _ = write!(out, "{}", other);
        }
    }
}

fn append_raw_template_parts(out: &mut String, parts: &[Value], subs: &[Value]) {
    let part_capacity: usize = parts
        .iter()
        .map(|value| match value {
            Value::String(text) => text.len(),
            _ => 8,
        })
        .sum();
    let subs_capacity: usize = subs
        .iter()
        .take(parts.len().saturating_sub(1))
        .map(|value| match value {
            Value::String(text) => text.len(),
            _ => 8,
        })
        .sum();
    out.reserve(part_capacity + subs_capacity);
    for (index, part) in parts.iter().enumerate() {
        push_string_value(out, part);
        if let Some(sub) = subs.get(index) {
            push_string_value(out, sub);
        }
    }
}

#[inline]
fn owned_string_arc(text: String) -> Arc<str> {
    crate::keys::owned_string_arc(text)
}

fn value_to_string_arc(value: &Value) -> Arc<str> {
    match value {
        Value::String(text) => Arc::clone(text),
        Value::I32(n) if *n >= 0 => crate::keys::small_index_arc(*n as usize)
            .unwrap_or_else(|| owned_string_arc(n.to_string())),
        Value::I32(n) => owned_string_arc(n.to_string()),
        Value::I64(n) if *n >= 0 => usize::try_from(*n)
            .ok()
            .and_then(crate::keys::small_index_arc)
            .unwrap_or_else(|| owned_string_arc(n.to_string())),
        Value::I64(n) => owned_string_arc(n.to_string()),
        Value::F32(n) => owned_string_arc(n.to_string()),
        Value::F64(n) => owned_string_arc(n.to_string()),
        Value::Bool(true) => crate::keys::string_arc("true"),
        Value::Bool(false) => crate::keys::string_arc("false"),
        Value::Null | Value::TypedNull(_) => crate::keys::string_arc("null"),
        Value::Undefined => crate::keys::string_arc("undefined"),
        Value::BigInt(n) => owned_string_arc(n.to_string()),
        other => owned_string_arc(crate::keys::value_display_string(other)),
    }
}

#[inline]
fn hex_nibble(byte: u8) -> Option<u8> {
    match byte {
        b'0'..=b'9' => Some(byte - b'0'),
        b'a'..=b'f' => Some(byte - b'a' + 10),
        b'A'..=b'F' => Some(byte - b'A' + 10),
        _ => None,
    }
}

#[inline]
fn char_value(ch: char) -> Value {
    crate::keys::char_value(ch)
}

#[inline]
pub fn uppercase_value(text: &str) -> Value {
    if text.is_ascii() {
        if !text.as_bytes().iter().any(u8::is_ascii_lowercase) {
            crate::keys::string_value(text)
        } else {
            s_owned(text.to_ascii_uppercase())
        }
    } else {
        s_owned(text.to_uppercase())
    }
}

#[inline]
pub fn lowercase_value(text: &str) -> Value {
    if text.is_ascii() {
        if !text.as_bytes().iter().any(u8::is_ascii_uppercase) {
            crate::keys::string_value(text)
        } else {
            s_owned(text.to_ascii_lowercase())
        }
    } else {
        s_owned(text.to_lowercase())
    }
}

pub fn shared_string_prototype() -> Value {
    Value::Object(
        STRING_PROTOTYPE
            .get_or_init(|| vybe_runtime::heap::alloc(Object::new()))
            .clone(),
    )
}

pub fn boxed_string(text: Arc<str>) -> Value {
    let mut obj = Object::new();
    let char_len = if text.is_ascii() {
        text.len()
    } else {
        text.chars().count()
    };
    obj.properties.reserve(char_len + 8);
    obj.properties
        .insert("__type".into(), crate::keys::string_value("String"));
    obj.properties
        .insert("__primitive".into(), Value::String(text.clone()));
    obj.properties
        .insert("__proto__".into(), shared_string_prototype());
    obj.properties
        .insert("length".into(), Value::I32(char_len as i32));

    let mut keys = Vec::with_capacity(char_len);
    if text.is_ascii() {
        for (index, byte) in text.as_bytes().iter().copied().enumerate() {
            let key = crate::keys::small_index_key(index)
                .map(str::to_owned)
                .unwrap_or_else(|| index.to_string());
            obj.properties
                .insert(key.clone(), crate::keys::char_value(byte as char));
            keys.push(crate::keys::string_value(key.as_str()));
        }
    } else {
        for (index, ch) in text.chars().enumerate() {
            let key = crate::keys::small_index_key(index)
                .map(str::to_owned)
                .unwrap_or_else(|| index.to_string());
            obj.properties.insert(key.clone(), char_value(ch));
            keys.push(crate::keys::string_value(key.as_str()));
        }
    }
    obj.properties.insert(
        "__keys".into(),
        Value::Object(vybe_runtime::heap::alloc(Object::new_array(keys))),
    );
    obj.properties.insert(
        "__nonenum".into(),
        Value::Object(vybe_runtime::heap::alloc(Object::new_array(vec![
            crate::keys::string_value("length"),
            crate::keys::string_value("toString"),
            crate::keys::string_value("valueOf"),
        ]))),
    );

    let arc = vybe_runtime::heap::alloc(obj);
    if let Value::Object(proto) = shared_string_prototype() {
        let proto = proto.lock().unwrap();
        for name in ["toString", "valueOf"] {
            if let Some(method) = proto.properties.get(name) {
                arc.lock()
                    .unwrap()
                    .properties
                    .insert(name.to_string(), method.clone());
            }
        }
    }
    Value::Object(arc)
}

fn to_string_primitive(ctx: &mut vybe_runtime::HostContext, value: Value) -> Arc<str> {
    if let Value::Object(obj) = &value {
        let function_name = {
            let o = obj.lock().unwrap();
            if matches!(
                o.kind,
                ObjectKind::Function(_) | ObjectKind::HostFunction(_)
            ) {
                o.properties.get("name").map(|name| match name {
                    Value::String(text) => text.as_ref().to_owned(),
                    other => value_to_string_arc(other).to_string(),
                })
            } else {
                None
            }
        };
        if let Some(name) = function_name {
            return if name.is_empty() {
                crate::keys::string_arc("function () { [native code] }")
            } else {
                crate::keys::concat3_arc("function ", &name, "() { [native code] }")
            };
        }
        let primitive = crate::value::to_primitive(ctx, &value, "string");
        return value_to_string_arc(&primitive);
    }
    value_to_string_arc(&value)
}

fn with_s_arg<R>(args: &[Value], idx: usize, f: impl FnOnce(&str) -> R) -> R {
    match args.get(idx) {
        Some(Value::String(text)) => f(text.as_ref()),
        Some(other) => {
            let text = value_to_string_arc(other);
            f(&text)
        }
        None => f(""),
    }
}

fn with_two_s_args<R>(
    args: &[Value],
    first: usize,
    second: usize,
    f: impl FnOnce(&str, &str) -> R,
) -> R {
    with_s_arg(args, first, |a| with_s_arg(args, second, |b| f(a, b)))
}

fn with_three_s_args<R>(
    args: &[Value],
    first: usize,
    second: usize,
    third: usize,
    f: impl FnOnce(&str, &str, &str) -> R,
) -> R {
    with_s_arg(args, first, |a| {
        with_s_arg(args, second, |b| with_s_arg(args, third, |c| f(a, b, c)))
    })
}

fn i32_arg(args: &[Value], idx: usize, default: i32) -> i32 {
    match args.get(idx) {
        Some(Value::I32(n)) => *n,
        Some(Value::F64(n)) => *n as i32,
        _ => default,
    }
}

fn to_integer_or_infinity(value: Option<&Value>, default_for_undefined: f64) -> f64 {
    let number = match value {
        None | Some(Value::Undefined) => default_for_undefined,
        Some(Value::Null) => 0.0,
        Some(Value::Bool(v)) => {
            if *v {
                1.0
            } else {
                0.0
            }
        }
        Some(Value::I32(n)) => *n as f64,
        Some(Value::I64(n)) => *n as f64,
        Some(Value::F32(n)) => *n as f64,
        Some(Value::F64(n)) => *n,
        Some(Value::String(text)) => {
            let trimmed = text.trim();
            if trimmed.is_empty() {
                0.0
            } else if trimmed == "Infinity" || trimmed == "+Infinity" {
                f64::INFINITY
            } else if trimmed == "-Infinity" {
                f64::NEG_INFINITY
            } else {
                trimmed.parse::<f64>().unwrap_or(f64::NAN)
            }
        }
        Some(other) => format!("{other}").parse::<f64>().unwrap_or(f64::NAN),
    };
    if number.is_nan() || number == 0.0 {
        0.0
    } else if number.is_infinite() {
        number
    } else {
        number.trunc()
    }
}

fn s_val(text: &str) -> Value {
    crate::keys::string_value(text)
}

fn s_owned(text: String) -> Value {
    crate::keys::owned_string_value(text)
}

fn utf16_units(text: &str) -> Vec<u16> {
    text.encode_utf16().collect()
}


// Two immutable primitive Strings cover the alternating receiver pattern in
// similarity loops. Weak keys keep their allocation identities stable without
// retaining the original Strings. At most 2x65536 UTF16 units are cached.
const CHAR_AT_CACHE_MAX_BYTES: usize = 65536;
#[derive(Default)]
struct CharAtView {
    source: Option<std::sync::Weak<str>>,
    ascii: bool,
    units: Vec<u16>,
}
#[derive(Default)]
struct CharAtViews {
    entries: [CharAtView; 2],
    recent: usize,
}
#[derive(Clone, Copy, Debug, PartialEq, Eq)]
enum CharAtUnit { Ascii(u8), Utf16(u16) }
impl CharAtViews {
    fn get(&mut self, source: &Arc<str>, index: usize) -> Option<CharAtUnit> {
        let found = self.entries.iter().position(|entry| {
            entry.source.as_ref().is_some_and(|old|
                std::ptr::eq(old.as_ptr(), Arc::as_ptr(source)))
        });
        let slot = found.unwrap_or_else(|| {
            let slot = self.entries.iter().position(|e| e.source.is_none())
                .unwrap_or(1-self.recent);
            let entry = &mut self.entries[slot];
            entry.ascii = source.is_ascii();
            entry.units.clear();
            if !entry.ascii {
                // UTF16 length never exceeds valid UTF8 byte length. Reserve
                // the full bounded upper limit exactly before iterator extend,
                // preventing Vec's geometric growth from crossing the budget.
                entry.units.reserve_exact(source.len());
                entry.units.extend(source.encode_utf16());
            }
            entry.source = Some(Arc::downgrade(source));
            slot
        });
        self.recent = slot;
        let entry=&self.entries[slot];
        if entry.ascii {
            source.as_bytes().get(index).copied().map(CharAtUnit::Ascii)
        } else {
            entry.units.get(index).copied().map(CharAtUnit::Utf16)
        }
    }
}

// Dynamic String.prototype dispatch reaches value.rs rather than these host
// exports. Keep its Unicode views bounded and independent of charAt's views.
// No user code runs while this thread-local buffer is borrowed.
thread_local! {
    static DYNAMIC_CHAR_CODE_VIEWS: std::cell::RefCell<CharAtViews> =
        std::cell::RefCell::new(CharAtViews::default());
}

pub(crate) fn dynamic_char_code_unit(source: &Arc<str>, index: usize) -> Option<u16> {
    if source.len() > CHAR_AT_CACHE_MAX_BYTES {
        // Preserve the existing uncached path for oversized inputs.
        return utf16_units(source).get(index).copied();
    }
    DYNAMIC_CHAR_CODE_VIEWS.with(|views| match views.borrow_mut().get(source, index) {
        Some(CharAtUnit::Ascii(unit)) => Some(u16::from(unit)),
        Some(CharAtUnit::Utf16(unit)) => Some(unit),
        None => None,
    })
}

// Search valid scalar strings without temporary UTF-16 buffers.
// Rust strings contain complete Unicode scalars. A valid scalar needle cannot
// match a low surrogate on its own, so a mid-pair start rounds to the next
// scalar boundary for nonempty needles. Empty needles retain the exact unit index.
pub(crate) fn utf8_index_of_units(source: &str, needle: &str, position: usize) -> Option<usize> {
    if position == 0 {
        // Search before classifying the whole source. An absent needle needs
        // no UTF-16 offset conversion; an early match only needs its prefix.
        let byte = source.find(needle)?;
        let prefix = &source[..byte];
        return Some(if prefix.is_ascii() {
            byte
        } else {
            prefix.encode_utf16().count()
        });
    }
    if source.is_ascii() && needle.is_ascii() {
        let start = position.min(source.len());
        return source[start..].find(needle).map(|offset| start + offset);
    }
    if needle.is_empty() {
        return Some(position.min(source.encode_utf16().count()));
    }
    let mut units = 0;
    let mut start_byte = source.len();
    for (byte, character) in source.char_indices() {
        if units >= position {
            start_byte = byte;
            break;
        }
        units += character.len_utf16();
    }
    let offset = source[start_byte..].find(needle)?;
    let match_byte = start_byte + offset;
    Some(units + source[start_byte..match_byte].encode_utf16().count())
}

/// Reverse search without materializing UTF-16 copies. A valid UTF-8 needle
/// starts and ends at scalar boundaries; its UTF-16 width is fixed. Thus a
/// match starts at/before `position` iff it ends at/before position+width.
pub(crate) fn utf8_last_index_of_units(source: &str, needle: &str, position: f64) -> Option<usize> {
    // A valid UTF-8 needle starts with an ASCII byte or a leading byte,
    // never a continuation byte. Byte matches therefore stay on scalar
    // boundaries, including when the library uses native SIMD search.
    if (position.is_infinite() && position.is_sign_positive())
        || position >= source.len() as f64 {
        let byte = memchr::memmem::rfind(source.as_bytes(), needle.as_bytes())?;
        return Some(vybe_runtime::utf16_units(&source[..byte]));
    }
    let position = if position <= 0.0 { 0 } else { position as usize };
    if needle.is_empty() {
        return Some(position.min(vybe_runtime::utf16_units(source)));
    }
    if needle.len() > source.len() { return None; }
    let limit = position.saturating_add(vybe_runtime::utf16_units(needle));
    let mut units = 0;
    let mut end_byte = 0;
    for (byte, ch) in source.char_indices() {
        let next = units + ch.len_utf16();
        if next > limit { break; }
        units = next;
        end_byte = byte + ch.len_utf8();
    }
    let byte = memchr::memmem::rfind(&source.as_bytes()[..end_byte], needle.as_bytes())?;
    // The boundary walk already counted this prefix. Subtract only the
    // matched suffix, instead of counting the whole prefix a second time.
    Some(units - vybe_runtime::utf16_units(&source[byte..end_byte]))
}

#[cfg(test)]
mod zero_position_search_tests {
    use super::utf8_index_of_units;

    #[test]
    fn search_offsets_match_utf16_oracle() {
        for source in ["", "plain ASCII", "aaaaab", "A\u{1f600}B\u{e9}C",
            "\u{1f600}\u{1f600}", "\0a\0"] {
            for needle in ["", "a", "plain", "\u{1f600}", "\u{e9}", "B",
                "\0", "missing", "\u{e9}C", "\u{1f600}B"] {
                let units: Vec<_> = source.encode_utf16().collect();
                let sought: Vec<_> = needle.encode_utf16().collect();
                for position in 0..=units.len() + 2 {
                    let expected = if sought.is_empty() {
                        Some(position.min(units.len()))
                    } else {
                        units.windows(sought.len()).enumerate()
                            .find(|(offset, window)| *offset >= position
                                && *window == sought.as_slice())
                            .map(|(offset, _)| offset)
                    };
                    assert_eq!(utf8_index_of_units(source, needle, position), expected,
                        "source={source:?}, needle={needle:?}, position={position}");
                }
            }
        }
    }
}


fn utf16_to_string(units: &[u16]) -> String {
    String::from_utf16_lossy(units)
}

#[cfg(test)]
pub(crate) mod char_code_at_cache_tests {
    use super::*;
    use vybe_runtime::{Chunk, Op};

    // Exact pre-cache charCodeAt body, retained solely as a test reference.
    pub(crate) fn before(args: &[Value]) -> Value {
        with_s_arg(args, 0, |s| {
            let pos = i32_arg(args, 1, 0);
            if pos < 0 { return Value::F64(f64::NAN); }
            if s.is_ascii() {
                return s.as_bytes().get(pos as usize)
                    .map(|byte| Value::F64(*byte as f64))
                    .unwrap_or(Value::F64(f64::NAN));
            }
            match utf16_units(s).get(pos as usize) {
                Some(unit) => Value::F64(*unit as f64),
                None => Value::F64(f64::NAN),
            }
        })
    }

    fn worker(old: bool) -> (VM, Value) {
        let mut owner = Chunk::new("char-code-owner");
        vybe_compiler::primitives::functions::emit_ref_func(&mut owner, 1, 0, 0);
        owner.emit_op(Op::RETURN, 0);
        let mut body = Chunk::new("char-code");
        body.arity = 2; body.param_count = 2; body.local_count = 2;
        body.emit_op_u16(Op::LOCAL_GET, 0, 0); body.emit_op_u16(Op::LOCAL_GET, 1, 0);
        let import = body.add_import("ecma:string", "charCodeAt");
        body.emit_call(import, 2, 0); body.emit_op(Op::RETURN, 0);
        let mut chunks = vec![owner, body];
        vybe_compiler::primitives::bundle::finalize_with_runtime_helpers(&mut chunks);
        let mut vm = VM::new();
        vybe_compiler::primitives::platforms::register_platforms_all(&mut vm);
        if old { vm.register_host_fn("ecma:string", "charCodeAt", Box::new(|_, args| before(args))); }
        let function = vm.run(chunks).unwrap();
        (vm, function)
    }

    fn same_number(a: Value, b: Value) {
        let (Value::F64(a), Value::F64(b)) = (a, b) else { panic!("expected numeric charCodeAt results") };
        assert!(a == b || (a.is_nan() && b.is_nan()), "{a} != {b}");
    }

    #[test]
    fn cached_char_code_at_preserves_utf16_nan_and_coercion_paths() {
        let (mut vm, f) = worker(false);
        let sources: Vec<Value> = ["", "A\0Ω中", "😀𐐀A", "é\u{301}"]
            .into_iter().map(|s| Value::String(Arc::from(s))).collect();
        for _ in 0..3 {
            for source in &sources {
                for pos in [-1, 0, 1, 2, 3, 4, 5, i32::MIN, i32::MAX] {
                    let args = [source.clone(), Value::I32(pos)];
                    same_number(vm.invoke(&f, &args).unwrap(), before(&args));
                    assert!(vm.last_exception.is_none());
                }
            }
        }
        // Independent UTF-16 oracle across every BMP scalar, in bounded blocks.
        for block in 0..256 {
            let text: String = (block * 256..(block + 1) * 256)
                .filter_map(char::from_u32).collect();
            let units: Vec<_> = text.encode_utf16().collect();
            let source = Value::String(Arc::from(text));
            for (index, unit) in units.iter().enumerate() {
                assert_eq!(vm.invoke(&f, &[source.clone(), Value::I32(index as i32)]).unwrap(),
                    Value::F64(*unit as f64));
            }
        }
        for length in [CHAR_AT_CACHE_MAX_BYTES - 1, CHAR_AT_CACHE_MAX_BYTES, CHAR_AT_CACHE_MAX_BYTES + 1] {
            let source = Value::String(Arc::from("A".repeat(length)));
            for pos in [0, length as i32 - 1, length as i32] {
                let args = [source.clone(), Value::I32(pos)];
                same_number(vm.invoke(&f, &args).unwrap(), before(&args));
            }
        }
        for receiver in [Value::String(Arc::from("ABC😀")), Value::Null, Value::Undefined, Value::I32(123), Value::Bool(true)] {
            for index in [Value::Undefined, Value::Null, Value::F64(f64::NAN), Value::F64(-0.0), Value::F64(1.8), Value::F64(f64::INFINITY), Value::F64(f64::NEG_INFINITY)] {
                let args = [receiver.clone(), index];
                same_number(vm.invoke(&f, &args).unwrap(), before(&args));
            }
        }
        let large = Value::String(Arc::from("Ω".repeat(CHAR_AT_CACHE_MAX_BYTES)));
        let args = [large, Value::I32(100)];
        same_number(vm.invoke(&f, &args).unwrap(), before(&args));
    }

    #[test]
    fn dynamic_charcode_view_bounds_roots_and_threads() {
        let check = || {
            DYNAMIC_CHAR_CODE_VIEWS.with(|views| *views.borrow_mut() = CharAtViews::default());
            let first: Arc<str> = Arc::from("Ω".repeat(32768));
            let second: Arc<str> = Arc::from("😀".repeat(16384));
            assert_eq!(dynamic_char_code_unit(&first, 32767), Some(937));
            assert_eq!(dynamic_char_code_unit(&second, 32767), Some(0xde00));
            let weak = Arc::downgrade(&first);
            assert_eq!(Arc::strong_count(&first), 1);
            drop(first);
            assert!(weak.upgrade().is_none());
            DYNAMIC_CHAR_CODE_VIEWS.with(|views| {
                let views = views.borrow();
                assert!(views.entries.iter().all(|e| e.units.capacity() <= CHAR_AT_CACHE_MAX_BYTES));
                assert_eq!(views.entries.iter().filter(|e| e.source.is_some()).count(), 2);
            });
            let oversized: Arc<str> = Arc::from("Ω".repeat(32769));
            assert_eq!(dynamic_char_code_unit(&oversized, 32768), Some(937));
            DYNAMIC_CHAR_CODE_VIEWS.with(|views| {
                let views = views.borrow();
                assert!(!views.entries.iter().any(|e| e.source.as_ref().is_some_and(|root|
                    std::ptr::eq(root.as_ptr(), Arc::as_ptr(&oversized)))));
            });
        };
        std::thread::spawn(check).join().unwrap();
        std::thread::spawn(check).join().unwrap();
    }

    #[test]
    #[ignore = "explicit opt0 adapter benchmark"]
    fn cached_char_code_at_debug_pairs() {
        assert!(cfg!(debug_assertions));
        let mut workers = [worker(true), worker(false)];
        let mut rows = Vec::new();
        for (kind, text, expected) in [("ascii", "A".repeat(8192), 65.0), ("unicode", "Ω".repeat(8192), 937.0)] {
            let args = [Value::String(Arc::from(text)), Value::I32(4096)];
            let batch = |worker: &mut (VM, Value), count: usize| {
                let start = std::time::Instant::now();
                for _ in 0..count { assert_eq!(worker.0.invoke(&worker.1, &args).unwrap(), Value::F64(expected)); }
                start.elapsed().as_nanos()
            };
            let mut counts = [1usize; 2];
            for side in 0..2 { while batch(&mut workers[side], counts[side]) < 500_000_000 { counts[side] *= 2; } }
            let mut samples = Vec::new(); let mut ratios = Vec::new();
            for pair in 0..7 {
                let mut ns = [0u128; 2];
                for side in if pair % 2 == 0 { [0, 1] } else { [1, 0] } {
                    ns[side] = batch(&mut workers[side], counts[side]);
                    samples.push((pair, side, counts[side], ns[side]));
                }
                ratios.push((ns[0] as f64 / counts[0] as f64) / (ns[1] as f64 / counts[1] as f64));
            }
            let mut sorted = ratios.clone(); sorted.sort_by(f64::total_cmp);
            rows.push((kind, sorted[3], ratios, samples));
            eprintln!("charCodeAt {kind}: {:.4}x; raw={:?}", sorted[3], rows.last().unwrap());
        }
    }


}

fn is_regexp_value(value: Option<&Value>) -> bool {
    let Some(Value::Object(obj)) = value else {
        return false;
    };
    let o = obj.lock().unwrap();
    matches!(o.properties.get("__type"), Some(Value::String(tag)) if tag.as_ref() == "RegExp")
}

/// Convert a possibly-negative ECMA-262 index into a clamped
/// non-negative position within `[0, len]`. Used by `slice` and
/// `at` which accept negative offsets.
fn clamp_signed(index: i32, len: usize) -> usize {
    let len_i = len as i32;
    let resolved = if index < 0 {
        (len_i + index).max(0)
    } else {
        index.min(len_i)
    };
    resolved.max(0) as usize
}

/// Declare an `ecma:string` function — same closure, plus the signature.
///
/// Like `ecma:array`, no resource binding: a string is a value, not a handle.
///
/// Only the handlers with a fixed parameter list are declared. §22.1.3 is full
/// of optional positions — `indexOf(search[, pos])`, `padStart(len[, filler])`,
/// `slice([start[, end]])` — and a Component Model signature cannot express
/// one, so declaring those would flag correct callers. `None` still means
/// UNKNOWN, so leaving them undeclared claims nothing. The same reasoning is
/// spelled out at length over `register` in `array.rs`.
fn string_fn(
    vm: &mut VM,
    name: &str,
    params: Vec<ValType>,
    results: Vec<ValType>,
    call: Box<dyn Fn(&mut HostContext, &[Value]) -> Value + Send + Sync>,
) {
    vm.register_host(
        HostFnDecl::new("ecma:string", name, call).with_sig(FuncSig {
            name: name.to_string(),
            params: Param::unnamed_list(params),
            results,
        }),
    );
}
/// Register a FREE FUNCTION — one whose type has no receiver parameter.
///
/// `encodeURIComponent(s)` takes a string and nothing else; it is not a method
/// being spelled as a call the way `padStart(s, 4)` is. Declaring that keeps
/// the shape identical whether it is called directly or handed to `map`.
fn string_fn_free(
    vm: &mut VM,
    name: &str,
    params: Vec<ValType>,
    results: Vec<ValType>,
    call: Box<dyn Fn(&mut HostContext, &[Value]) -> Value + Send + Sync>,
) {
    vm.register_host(
        HostFnDecl::new("ecma:string", name, call)
            .with_sig(FuncSig {
                name: name.to_string(),
                params: Param::unnamed_list(params),
                results,
            })
            .without_receiver(),
    );
}

/// The string operand — every §22.1.3 method takes its receiver first.
fn str_t() -> ValType {
    ValType::String
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
    register_iterator(vm);
}

// ── Adapter convenience methods ──────────────────────────────────
//
// Not in ECMA-262 §22.1 but ubiquitous across Python str /
// Ruby String / .NET — the predicates and case-conversions every
// language exposes. Live here as one-line Rust impls so the cross-
// language profile entries can target a single host fn surface (the
// `ecma:string` namespace). Same precedent as the ecma:array
// adapters (`clear`, `first`, `last`, `removeAt`, etc.).

fn register_adapters(vm: &mut VM) {
    // Python str.isdigit() — true iff non-empty and every char is
    // a Unicode decimal digit.
    string_fn(
        vm,
        "isdigit",
        vec![str_t()],
        vec![ValType::Bool],
        Box::new(|_ctx, args| {
            with_s_arg(args, 0, |s| {
                Value::Bool(!s.is_empty() && s.chars().all(|c| c.is_ascii_digit()))
            })
        }),
    );
    string_fn(
        vm,
        "isalpha",
        vec![str_t()],
        vec![ValType::Bool],
        Box::new(|_ctx, args| {
            with_s_arg(args, 0, |s| {
                Value::Bool(!s.is_empty() && s.chars().all(|c| c.is_alphabetic()))
            })
        }),
    );
    string_fn(
        vm,
        "isalnum",
        vec![str_t()],
        vec![ValType::Bool],
        Box::new(|_ctx, args| {
            with_s_arg(args, 0, |s| {
                Value::Bool(!s.is_empty() && s.chars().all(|c| c.is_alphanumeric()))
            })
        }),
    );
    string_fn(
        vm,
        "isspace",
        vec![str_t()],
        vec![ValType::Bool],
        Box::new(|_ctx, args| {
            with_s_arg(args, 0, |s| {
                Value::Bool(!s.is_empty() && s.chars().all(|c| c.is_whitespace()))
            })
        }),
    );
    string_fn(
        vm,
        "isupper",
        vec![str_t()],
        vec![ValType::Bool],
        Box::new(|_ctx, args| {
            with_s_arg(args, 0, |s| {
                let mut has_cased = false;
                for c in s.chars() {
                    if c.is_uppercase() {
                        has_cased = true;
                    } else if c.is_lowercase() {
                        return Value::Bool(false);
                    }
                }
                Value::Bool(has_cased)
            })
        }),
    );
    string_fn(
        vm,
        "islower",
        vec![str_t()],
        vec![ValType::Bool],
        Box::new(|_ctx, args| {
            with_s_arg(args, 0, |s| {
                let mut has_cased = false;
                for c in s.chars() {
                    if c.is_lowercase() {
                        has_cased = true;
                    } else if c.is_uppercase() {
                        return Value::Bool(false);
                    }
                }
                Value::Bool(has_cased)
            })
        }),
    );

    // Python str.title() / Ruby capitalize-each-word — uppercase the
    // first letter of each word, lowercase the rest.
    string_fn(
        vm,
        "title",
        vec![str_t()],
        vec![str_t()],
        Box::new(|_ctx, args| {
            with_s_arg(args, 0, |s| {
                let mut out = String::with_capacity(s.len());
                let mut prev_alnum = false;
                for c in s.chars() {
                    if c.is_alphanumeric() {
                        if prev_alnum {
                            for lc in c.to_lowercase() {
                                out.push(lc);
                            }
                        } else {
                            for uc in c.to_uppercase() {
                                out.push(uc);
                            }
                        }
                        prev_alnum = true;
                    } else {
                        out.push(c);
                        prev_alnum = false;
                    }
                }
                s_owned(out)
            })
        }),
    );

    // Python str.swapcase() — uppercase chars become lowercase and vice
    // versa; non-cased chars unchanged.
    string_fn(
        vm,
        "swapcase",
        vec![str_t()],
        vec![str_t()],
        Box::new(|_ctx, args| {
            with_s_arg(args, 0, |s| {
                let mut out = String::with_capacity(s.len());
                for c in s.chars() {
                    if c.is_uppercase() {
                        for lc in c.to_lowercase() {
                            out.push(lc);
                        }
                    } else if c.is_lowercase() {
                        for uc in c.to_uppercase() {
                            out.push(uc);
                        }
                    } else {
                        out.push(c);
                    }
                }
                s_owned(out)
            })
        }),
    );

    // Ruby `s.tr(from, to)` — translate chars in `from` to the parallel
    // char in `to`. Chars in `from` not in `to` are left unchanged.
    // Simplified: doesn't handle ranges (`a-z`) or negation (`^abc`) —
    // those need Ruby-specific intrinsics.
    string_fn(
        vm,
        "tr",
        vec![str_t(), str_t(), str_t()],
        vec![str_t()],
        Box::new(|_ctx, args| {
            with_three_s_args(args, 0, 1, 2, |s, from_arg, to_arg| {
                if s.is_ascii() && from_arg.is_ascii() && to_arg.is_ascii() {
                    let mut table = [0u8; 256];
                    for (index, slot) in table.iter_mut().enumerate() {
                        *slot = index as u8;
                    }
                    let mut seen = [false; 256];
                    for (from, to) in from_arg.bytes().zip(to_arg.bytes()) {
                        let index = from as usize;
                        if !seen[index] {
                            table[index] = to;
                            seen[index] = true;
                        }
                    }
                    let mut out = Vec::with_capacity(s.len());
                    for byte in s.bytes() {
                        out.push(table[byte as usize]);
                    }
                    return s_val(std::str::from_utf8(&out).unwrap_or(""));
                }
                let from: Vec<char> = from_arg.chars().collect();
                let to: Vec<char> = to_arg.chars().collect();
                let out: String = s
                    .chars()
                    .map(|c| {
                        from.iter()
                            .position(|&fc| fc == c)
                            .and_then(|i| to.get(i).copied())
                            .unwrap_or(c)
                    })
                    .collect();
                s_owned(out)
            })
        }),
    );
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
    vm.register_free_fn(
        "ecma:string",
        "String",
        Box::new(|ctx, args| {
            let v = args.first().cloned().unwrap_or(Value::Undefined);
            Value::String(to_string_primitive(ctx, v))
        }),
    );
    vm.register_free_fn(
        "ecma:string",
        "new",
        Box::new(|ctx, args| {
            let v = args.first().cloned().unwrap_or(Value::Undefined);
            boxed_string(to_string_primitive(ctx, v))
        }),
    );
}

// ── Query ops (length, char access) ───────────────────────────────

fn register_query_ops(vm: &mut VM) {
    // String.prototype.length — wasm:js-string already covers this
    // via a host fn. Re-register under ecma:string for callers that
    // want the canonical JS-runtime name.
    // §22.1.4.1 `length` counts UTF-16 CODE UNITS; the handler answers F64, so
    // the declaration says F64 rather than rounding the contract to i32.
    string_fn(
        vm,
        "length",
        vec![str_t()],
        vec![ValType::F64],
        Box::new(|_ctx, args| {
            with_s_arg(args, 0, |text| {
                Value::F64(text.encode_utf16().count() as f64)
            })
        }),
    );

    // String.prototype.charAt(pos). Keep the non-String coercion path exact;
    // only the immutable primitive String producer gets an indexed view.
    let char_at_views = Arc::new(Mutex::new(CharAtViews::default()));
    // Separate bounded views prevent charCodeAt on a third receiver from
    // evicting the two strings used by an interleaved charAt loop.
    let char_code_at_views = Arc::new(Mutex::new(CharAtViews::default()));
    vm.register_host_fn(
        "ecma:string",
        "charAt",
        Box::new(move |_ctx, args| {
            if let Some(Value::String(source)) = args.first() {
                let pos = i32_arg(args, 1, 0);
                if pos < 0 { return s_val(""); }
                if source.len() <= CHAR_AT_CACHE_MAX_BYTES {
                    let unit = char_at_views.lock().unwrap().get(source, pos as usize);
                    // Release private cache lock before entering the key pool.
                    return match unit {
                        Some(CharAtUnit::Ascii(byte)) => crate::keys::char_value(byte as char),
                        Some(CharAtUnit::Utf16(unit)) => s_val(&String::from_utf16_lossy(&[unit])),
                        None => s_val(""),
                    };
                }
            }
            with_s_arg(args, 0, |s| {
                let pos = i32_arg(args, 1, 0);
                if pos < 0 {
                    return s_val("");
                }
                if s.is_ascii() {
                    let index = pos as usize;
                    return match s.as_bytes().get(index) {
                        Some(byte) => crate::keys::char_value(*byte as char),
                        None => s_val(""),
                    };
                }
                // §22.1.3.2: one UTF-16 code unit, not a code point — an
                // unpaired surrogate half surfaces as U+FFFD (UTF-8 storage).
                match utf16_units(s).get(pos as usize) {
                    Some(unit) => s_val(&String::from_utf16_lossy(&[*unit])),
                    None => s_val(""),
                }
            })
        }),
    );

    // String.prototype.charCodeAt(pos)
    vm.register_host_fn(
        "ecma:string",
        "charCodeAt",
        Box::new(move |_ctx, args| {
            if let Some(Value::String(source)) = args.first() {
                if source.len() <= CHAR_AT_CACHE_MAX_BYTES {
                    let pos = i32_arg(args, 1, 0);
                    if pos < 0 { return Value::F64(f64::NAN); }
                    return match char_code_at_views.lock().unwrap().get(source, pos as usize) {
                        Some(CharAtUnit::Ascii(byte)) => Value::F64(byte as f64),
                        Some(CharAtUnit::Utf16(unit)) => Value::F64(unit as f64),
                        None => Value::F64(f64::NAN),
                    };
                }
            }
            with_s_arg(args, 0, |s| {
                let pos = i32_arg(args, 1, 0);
                if pos < 0 {
                    return Value::F64(f64::NAN);
                }
                if s.is_ascii() {
                    return s
                        .as_bytes()
                        .get(pos as usize)
                        .map(|byte| Value::F64(*byte as f64))
                        .unwrap_or(Value::F64(f64::NAN));
                }
                match utf16_units(s).get(pos as usize) {
                    Some(unit) => Value::F64(*unit as f64),
                    None => Value::F64(f64::NAN),
                }
            })
        }),
    );

    // String.prototype.codePointAt(pos)
    vm.register_host_fn(
        "ecma:string",
        "codePointAt",
        Box::new(|_ctx, args| {
            with_s_arg(args, 0, |s| {
                let pos = i32_arg(args, 1, 0);
                if pos < 0 {
                    return Value::Undefined;
                }
                if s.is_ascii() {
                    return s
                        .as_bytes()
                        .get(pos as usize)
                        .map(|byte| Value::F64(*byte as f64))
                        .unwrap_or(Value::Undefined);
                }
                match s.chars().nth(pos as usize) {
                    Some(ch) => Value::F64(ch as u32 as f64),
                    None => Value::Undefined,
                }
            })
        }),
    );

    // String.prototype.at(index) — ES2022; supports negative offsets.
    vm.register_host_fn(
        "ecma:string",
        "at",
        Box::new(|_ctx, args| {
            with_s_arg(args, 0, |s| {
                let index = i32_arg(args, 1, 0);
                if s.is_ascii() {
                    let len_i = s.len() as i32;
                    let resolved = if index < 0 { len_i + index } else { index };
                    if resolved < 0 || resolved >= len_i {
                        return Value::Undefined;
                    }
                    let byte = s.as_bytes()[resolved as usize];
                    return crate::keys::char_value(byte as char);
                }
                let len_i = s.chars().count() as i32;
                let resolved = if index < 0 { len_i + index } else { index };
                if resolved < 0 || resolved >= len_i {
                    return Value::Undefined;
                }
                match s.chars().nth(resolved as usize) {
                    Some(ch) => char_value(ch),
                    None => Value::Undefined,
                }
            })
        }),
    );

    string_fn(
        vm,
        "toString",
        vec![str_t()],
        vec![ValType::String],
        Box::new(|ctx, args| {
            let value = args
                .first()
                .cloned()
                .unwrap_or_else(|| crate::keys::string_value(""));
            Value::String(to_string_primitive(ctx, value))
        }),
    );

    // String.prototype.valueOf — returns the primitive string itself.
    string_fn(
        vm,
        "valueOf",
        vec![str_t()],
        vec![ValType::String],
        Box::new(|ctx, args| {
            let value = args
                .first()
                .cloned()
                .unwrap_or_else(|| crate::keys::string_value(""));
            Value::String(to_string_primitive(ctx, value))
        }),
    );
}

// ── Extract ops (substring, slice, concat) ────────────────────────

fn register_extract_ops(vm: &mut VM) {
    // String.prototype.concat(...strings)
    vm.register_host_fn(
        "ecma:string",
        "concat",
        Box::new(|_ctx, args| {
            let capacity = args
                .iter()
                .map(|value| match value {
                    Value::String(text) => text.len(),
                    _ => 0,
                })
                .sum();
            let mut out = String::with_capacity(capacity);
            for a in args {
                match a {
                    Value::String(text) => out.push_str(text),
                    other => {
                        let _ = write!(out, "{}", other);
                    }
                }
            }
            s_owned(out)
        }),
    );

    // String.prototype.substring(indexStart, indexEnd?) — clamps
    // both args to [0, len], swaps if start > end.
    vm.register_host_fn(
        "ecma:string",
        "substring",
        Box::new(|_ctx, args| {
            with_s_arg(args, 0, |s| {
                if s.is_ascii() {
                    let len = s.len();
                    let start_raw = i32_arg(args, 1, 0).max(0) as usize;
                    let start = start_raw.min(len);
                    let end = if args.len() >= 3 {
                        (i32_arg(args, 2, len as i32).max(0) as usize).min(len)
                    } else {
                        len
                    };
                    let (lo, hi) = if start <= end {
                        (start, end)
                    } else {
                        (end, start)
                    };
                    return s_val(&s[lo..hi]);
                }
                let chars: Vec<char> = s.chars().collect();
                let len = chars.len();
                let start_raw = i32_arg(args, 1, 0).max(0) as usize;
                let start = start_raw.min(len);
                let end = if args.len() >= 3 {
                    (i32_arg(args, 2, len as i32).max(0) as usize).min(len)
                } else {
                    len
                };
                let (lo, hi) = if start <= end {
                    (start, end)
                } else {
                    (end, start)
                };
                s_owned(chars[lo..hi].iter().collect::<String>())
            })
        }),
    );

    // String.prototype.substr(start, length?) — legacy Annex B method.
    vm.register_host_fn(
        "ecma:string",
        "substr",
        Box::new(|_ctx, args| {
            with_s_arg(args, 0, |s| {
                if s.is_ascii() {
                    let len = s.len() as i32;
                    let start_raw = i32_arg(args, 1, 0);
                    let start = if start_raw < 0 {
                        (len + start_raw).max(0)
                    } else {
                        start_raw.min(len)
                    } as usize;
                    let end = match args.get(2) {
                        Some(Value::Undefined) | Some(Value::Null) | None => s.len(),
                        Some(_) => {
                            let count = i32_arg(args, 2, 0).max(0) as usize;
                            start.saturating_add(count).min(s.len())
                        }
                    };
                    return s_val(&s[start..end]);
                }
                let units = utf16_units(s);
                let len = units.len() as i32;
                let start_raw = i32_arg(args, 1, 0);
                let start = if start_raw < 0 {
                    (len + start_raw).max(0)
                } else {
                    start_raw.min(len)
                } as usize;
                let end = match args.get(2) {
                    Some(Value::Undefined) | Some(Value::Null) | None => units.len(),
                    Some(_) => {
                        let count = i32_arg(args, 2, 0).max(0) as usize;
                        start.saturating_add(count).min(units.len())
                    }
                };
                s_owned(utf16_to_string(&units[start..end]))
            })
        }),
    );

    // String.prototype.slice(start?, end?) — supports negative
    // offsets (count from end). Differs from substring: no swap on
    // out-of-order args (returns empty instead).
    vm.register_host_fn(
        "ecma:string",
        "slice",
        Box::new(|_ctx, args| {
            with_s_arg(args, 0, |s| {
                // §22.1.3.21: indices are UTF-16 code units (same unit space
                // as `length`/`charCodeAt`), not code points.
                if s.is_ascii() {
                    let len = s.len();
                    let start = clamp_signed(i32_arg(args, 1, 0), len);
                    let end = if args.len() >= 3 {
                        clamp_signed(i32_arg(args, 2, len as i32), len)
                    } else {
                        len
                    };
                    return if start >= end {
                        s_val("")
                    } else {
                        s_val(&s[start..end])
                    };
                }
                let units = utf16_units(s);
                let len = units.len();
                let start = clamp_signed(i32_arg(args, 1, 0), len);
                let end = if args.len() >= 3 {
                    clamp_signed(i32_arg(args, 2, len as i32), len)
                } else {
                    len
                };
                if start >= end {
                    return s_val("");
                }
                s_owned(String::from_utf16_lossy(&units[start..end]))
            })
        }),
    );
}

// ── Casing ops ────────────────────────────────────────────────────

fn register_casing_ops(vm: &mut VM) {
    string_fn(
        vm,
        "toUpperCase",
        vec![str_t()],
        vec![str_t()],
        Box::new(|_ctx, args| with_s_arg(args, 0, uppercase_value)),
    );
    string_fn(
        vm,
        "toLowerCase",
        vec![str_t()],
        vec![str_t()],
        Box::new(|_ctx, args| with_s_arg(args, 0, lowercase_value)),
    );
    // The locale-aware variants use the same Rust impl until
    // locale data is wired in; spec says the result is implementation
    // defined for non-locale data anyway.
    string_fn(
        vm,
        "toLocaleUpperCase",
        vec![str_t()],
        vec![str_t()],
        Box::new(|_ctx, args| with_s_arg(args, 0, uppercase_value)),
    );
    string_fn(
        vm,
        "toLocaleLowerCase",
        vec![str_t()],
        vec![str_t()],
        Box::new(|_ctx, args| with_s_arg(args, 0, lowercase_value)),
    );
}

// ── Trim ops ──────────────────────────────────────────────────────

fn register_trim_ops(vm: &mut VM) {
    string_fn(
        vm,
        "trim",
        vec![str_t()],
        vec![str_t()],
        Box::new(|_ctx, args| with_s_arg(args, 0, |text| s_val(text.trim()))),
    );
    string_fn(
        vm,
        "trimStart",
        vec![str_t()],
        vec![str_t()],
        Box::new(|_ctx, args| with_s_arg(args, 0, |text| s_val(text.trim_start()))),
    );
    string_fn(
        vm,
        "trimEnd",
        vec![str_t()],
        vec![str_t()],
        Box::new(|_ctx, args| with_s_arg(args, 0, |text| s_val(text.trim_end()))),
    );
}

// ── Pad ops ───────────────────────────────────────────────────────

fn register_pad_ops(vm: &mut VM) {
    fn pad(args: &[Value], at_start: bool) -> Value {
        with_s_arg(args, 0, |s| {
            // §22.1.3.17.1 StringPad: maxLength and the filler truncation are
            // measured in UTF-16 CODE UNITS ("1".padEnd(3,"🌟") → "1🌟": the
            // two-unit star fills exactly). A truncation that splits a
            // surrogate pair would produce a lone surrogate per spec — our
            // UTF-8 backing substitutes U+FFFD at that edge.
            let target = i32_arg(args, 1, 0).max(0) as usize;
            let ascii_source_len = if s.is_ascii() { Some(s.len()) } else { None };
            let source_len = ascii_source_len.unwrap_or_else(|| s.encode_utf16().count());
            if source_len >= target {
                return s_val(s);
            }
            with_s_arg(args, 2, |pad_str| {
                let pad_str = if args.len() >= 3 { pad_str } else { " " };
                if pad_str.is_empty() {
                    return s_val(s);
                }
                if s.is_ascii() && pad_str.is_ascii() {
                    let needed = target - source_len;
                    let mut filler = String::with_capacity(needed);
                    let full_repeats = needed / pad_str.len();
                    for _ in 0..full_repeats {
                        filler.push_str(pad_str);
                    }
                    let remainder = needed % pad_str.len();
                    if remainder != 0 {
                        filler.push_str(&pad_str[..remainder]);
                    }
                    let mut result = String::with_capacity(filler.len() + s.len());
                    if at_start {
                        result.push_str(&filler);
                        result.push_str(s);
                    } else {
                        result.push_str(s);
                        result.push_str(&filler);
                    }
                    return s_owned(result);
                }
                let units: Vec<u16> = s.encode_utf16().collect();
                let pad_units: Vec<u16> = pad_str.encode_utf16().collect();
                let needed = target - units.len();
                let mut filler_units: Vec<u16> = Vec::with_capacity(needed);
                for i in 0..needed {
                    filler_units.push(pad_units[i % pad_units.len()]);
                }
                let filler = String::from_utf16_lossy(&filler_units);
                let mut result = String::with_capacity(filler.len() + s.len());
                if at_start {
                    result.push_str(&filler);
                    result.push_str(s);
                } else {
                    result.push_str(s);
                    result.push_str(&filler);
                }
                s_owned(result)
            })
        })
    }
    vm.register_host_fn(
        "ecma:string",
        "padStart",
        Box::new(|_ctx, args| pad(args, true)),
    );
    vm.register_host_fn(
        "ecma:string",
        "padEnd",
        Box::new(|_ctx, args| pad(args, false)),
    );
}

// ── Search ops ────────────────────────────────────────────────────

#[cfg(feature = "primitive-benchmarks")]
fn legacy_last_index_of_units(s: &str, needle: &str, position: f64) -> Value {
                let hay = utf16_units(s);
                let ndl = utf16_units(needle);
                let pos = position;
                let end = if pos.is_infinite() && pos.is_sign_positive() {
                    hay.len()
                } else if pos <= 0.0 {
                    0
                } else if pos >= hay.len() as f64 {
                    hay.len()
                } else {
                    pos as usize
                };
                if ndl.is_empty() {
                    return Value::F64(end as f64);
                }
                if ndl.len() > hay.len() {
                    return Value::F64(-1.0);
                }
                let start = end.min(hay.len() - ndl.len());
                for i in (0..=start).rev() {
                    if hay[i..i + ndl.len()] == ndl[..] {
                        return Value::F64(i as f64);
                    }
                }
                Value::F64(-1.0)
}

/// Host-side A/B registration. The guest import and signature stay identical.
#[cfg(feature = "primitive-benchmarks")]
pub fn benchmark_last_index_of(vm: &mut VM, fast: bool) {
    register_last_index_of(vm, fast);
}

fn register_last_index_of(vm: &mut VM, #[cfg(feature = "primitive-benchmarks")] fast: bool) {
    vm.register_host_fn(
        "ecma:string",
        "lastIndexOf",
        Box::new(move |_ctx, args| {
            // ECMA-262 §22.1.3.10: UTF-16 code-unit index of the last match.
            with_two_s_args(args, 0, 1, |s, needle| {
                if s.is_ascii() && needle.is_ascii() {
                    let pos = to_integer_or_infinity(args.get(2), f64::INFINITY);
                    let end = if pos.is_infinite() && pos.is_sign_positive() {
                        s.len()
                    } else if pos <= 0.0 {
                        0
                    } else if pos >= s.len() as f64 {
                        s.len()
                    } else {
                        pos as usize
                    };
                    if needle.is_empty() {
                        return Value::F64(end as f64);
                    }
                    if needle.len() > s.len() {
                        return Value::F64(-1.0);
                    }
                    let search_end = end.min(s.len().saturating_sub(needle.len()));
                    return match s[..search_end + needle.len()].rfind(needle) {
                        Some(byte_idx) => Value::F64(byte_idx as f64),
                        None => Value::F64(-1.0),
                    };
                }
                let position = to_integer_or_infinity(args.get(2), f64::INFINITY);
                #[cfg(feature = "primitive-benchmarks")]
                if !fast { return legacy_last_index_of_units(s, needle, position); }
                Value::F64(utf8_last_index_of_units(s, needle, position)
                    .map(|index| index as f64).unwrap_or(-1.0))
            })
        }),
    );

}

/// First complete scalar boundary at/after a UTF16 position, clamped at EOF.
/// `exact` is false only when the position splits a surrogate pair. Search
/// can skip to that boundary; a nonempty prefix/suffix cannot match across it.
fn utf8_clamped_boundary(source: &str, position: usize) -> (usize, bool) {
    let mut units = 0;
    for (byte, ch) in source.char_indices() {
        if units >= position { return (byte, units == position); }
        units += ch.len_utf16();
    }
    (source.len(), units <= position)
}

/// Host-only controls with identical import names and conversion paths.
#[cfg(feature = "primitive-benchmarks")]
pub fn benchmark_utf8_boundaries(vm: &mut VM, fast: bool) {
    register_includes(vm, fast);
    register_starts_with(vm, fast);
    register_ends_with(vm, fast);
}

#[cfg(feature = "primitive-benchmarks")]
fn legacy_includes(s: &str, needle: &str, args: &[Value]) -> Value {
    let pos = i32_arg(args, 2, 0).max(0) as usize;
    if pos == 0 {
        return Value::Bool(s.contains(needle));
    }
    if s.is_ascii() && needle.is_ascii() {
        let start = pos.min(s.len());
        return Value::Bool(s[start..].contains(needle));
    }
    if needle.is_empty() {
        return Value::Bool(true);
    }
    let hay_units = utf16_units(s);
    let needle_units = utf16_units(needle);
    let start = pos.min(hay_units.len());
    Value::Bool(
        hay_units[start..]
            .windows(needle_units.len())
            .any(|window| window == needle_units.as_slice()),
    )
}

fn register_includes(vm: &mut VM, #[cfg(feature = "primitive-benchmarks")] fast: bool) {
    vm.register_host_fn(
        "ecma:string",
        "includes",
        Box::new(move |ctx, args| {
            if is_regexp_value(args.get(1)) {
                ctx.throw_value(crate::error::new_error(
                    ctx,
                    "TypeError",
                    "First argument to String.prototype.includes must not be a RegExp",
                ));
                return Value::Null;
            }
            with_two_s_args(args, 0, 1, |s, needle| {
                #[cfg(feature = "primitive-benchmarks")]
                if !fast { return legacy_includes(s, needle, args); }
                let pos = i32_arg(args, 2, 0).max(0) as usize;
                if pos == 0 { return Value::Bool(s.contains(needle)); }
                if needle.is_empty() { return Value::Bool(true); }
                if s.is_ascii() && needle.is_ascii() {
                    let start = pos.min(s.len());
                    return Value::Bool(s[start..].contains(needle));
                }
                let (start, _) = utf8_clamped_boundary(s, pos);
                Value::Bool(memchr::memmem::find(&s.as_bytes()[start..], needle.as_bytes()).is_some())
            })
        }),
    );
}

#[cfg(feature = "primitive-benchmarks")]
fn legacy_starts_with(s: &str, needle: &str, args: &[Value]) -> Value {
    if s.is_ascii() && needle.is_ascii() {
        let start = (i32_arg(args, 2, 0).max(0) as usize).min(s.len());
        return Value::Bool(s[start..].starts_with(needle));
    }
    let hay = utf16_units(s);
    let ndl = utf16_units(needle);
    let start = (i32_arg(args, 2, 0).max(0) as usize).min(hay.len());
    if start + ndl.len() > hay.len() {
        return Value::Bool(false);
    }
    Value::Bool(hay[start..start + ndl.len()] == ndl[..])
}

fn register_starts_with(vm: &mut VM, #[cfg(feature = "primitive-benchmarks")] fast: bool) {
    vm.register_host_fn(
        "ecma:string",
        "startsWith",
        Box::new(move |_ctx, args| {
            // ECMA-262 §22.1.3.23: `position` is a UTF-16 code-unit index.
            with_two_s_args(args, 0, 1, |s, needle| {
                #[cfg(feature = "primitive-benchmarks")]
                if !fast { return legacy_starts_with(s, needle, args); }
                let pos = i32_arg(args, 2, 0).max(0) as usize;
                if pos == 0 { return Value::Bool(s.starts_with(needle)); }
                if needle.is_empty() { return Value::Bool(true); }
                if s.is_ascii() && needle.is_ascii() {
                    return Value::Bool(s[pos.min(s.len())..].starts_with(needle));
                }
                let (start, exact) = utf8_clamped_boundary(s, pos);
                Value::Bool(exact && s[start..].starts_with(needle))
            })
        }),
    );
}

#[cfg(feature = "primitive-benchmarks")]
fn legacy_ends_with(s: &str, needle: &str, args: &[Value]) -> Value {
    if s.is_ascii() && needle.is_ascii() {
        let end = if args.len() >= 3 {
            (i32_arg(args, 2, s.len() as i32).max(0) as usize).min(s.len())
        } else {
            s.len()
        };
        return Value::Bool(s[..end].ends_with(needle));
    }
    let hay = utf16_units(s);
    let ndl = utf16_units(needle);
    let end = if args.len() >= 3 {
        (i32_arg(args, 2, hay.len() as i32).max(0) as usize).min(hay.len())
    } else {
        hay.len()
    };
    if ndl.len() > end {
        return Value::Bool(false);
    }
    Value::Bool(hay[end - ndl.len()..end] == ndl[..])
}

fn register_ends_with(vm: &mut VM, #[cfg(feature = "primitive-benchmarks")] fast: bool) {
    vm.register_host_fn(
        "ecma:string",
        "endsWith",
        Box::new(move |_ctx, args| {
            // ECMA-262 §22.1.3.6: `endPosition` is a UTF-16 code-unit index.
            with_two_s_args(args, 0, 1, |s, needle| {
                #[cfg(feature = "primitive-benchmarks")]
                if !fast { return legacy_ends_with(s, needle, args); }
                if needle.is_empty() { return Value::Bool(true); }
                let position = match args.get(2) {
                    None => return Value::Bool(s.ends_with(needle)),
                    Some(Value::I32(n)) => *n,
                    Some(Value::F64(n)) => *n as i32,
                    // i32_arg historically uses the source length as the
                    // default for every other provided value. Its signed cast
                    // only differs from the full end on exceptionally large
                    // strings, so retain that conversion when necessary.
                    Some(_) if s.len() <= i32::MAX as usize => return Value::Bool(s.ends_with(needle)),
                    Some(_) => vybe_runtime::utf16_units(s) as i32,
                }.max(0) as usize;
                if s.is_ascii() && needle.is_ascii() {
                    return Value::Bool(s[..position.min(s.len())].ends_with(needle));
                }
                let (end, exact) = utf8_clamped_boundary(s, position);
                Value::Bool(exact && s[..end].ends_with(needle))
            })
        }),
    );
}

fn register_search_ops(vm: &mut VM) {
    register_includes(vm, #[cfg(feature = "primitive-benchmarks")] true);

    vm.register_host_fn(
        "ecma:string",
        "indexOf",
        Box::new(|_ctx, args| {
            // Preserve receiver/needle coercion order and existing position
            // conversion. The search reports UTF-16 units, not UTF-8 bytes.
            with_two_s_args(args, 0, 1, |s, needle| {
                let pos = i32_arg(args, 2, 0).max(0) as usize;
                Value::F64(utf8_index_of_units(s, needle, pos)
                    .map(|index| index as f64).unwrap_or(-1.0))
            })
        }),
    );

    register_last_index_of(vm, #[cfg(feature = "primitive-benchmarks")] true);

    register_starts_with(vm, #[cfg(feature = "primitive-benchmarks")] true);

    register_ends_with(vm, #[cfg(feature = "primitive-benchmarks")] true);
}

// ── Modify ops (repeat, replace, replaceAll) ──────────────────────

fn register_modify_ops(vm: &mut VM) {
    vm.register_host_fn(
        "ecma:string",
        "repeat",
        Box::new(|ctx, args| {
            with_s_arg(args, 0, |s| {
                // §22.1.3.19 steps 3–4: count < 0 or count = +∞ throws
                // RangeError (NaN → 0 via ToIntegerOrInfinity).
                let n = args.get(1).map(|v| v.as_f64()).unwrap_or(0.0);
                if n < 0.0 || (n.is_infinite() && n > 0.0) {
                    ctx.throw_value(crate::error::new_error(
                        ctx,
                        "RangeError",
                        "Invalid count value",
                    ));
                    return Value::Undefined;
                }
                let n = if n.is_nan() { 0 } else { n as usize };
                match n {
                    0 => s_val(""),
                    1 => s_val(s),
                    _ => s_owned(s.repeat(n)),
                }
            })
        }),
    );

    // ECMA-262 §22.1.3.18: replace with a string searchValue
    // replaces only the FIRST match. Rust's `str::replacen(.., 1)`
    // matches that.
    string_fn(
        vm,
        "replace",
        vec![str_t(), str_t(), str_t()],
        vec![str_t()],
        Box::new(|_ctx, args| {
            with_three_s_args(args, 0, 1, 2, |s, search, replace| {
                match s.find(search) {
                    None => s_val(s),
                    Some(start) => {
                        let capacity = s.len().saturating_sub(search.len()).saturating_add(replace.len());
                        let mut out = String::with_capacity(capacity);
                        out.push_str(&s[..start]);
                        out.push_str(replace);
                        out.push_str(&s[start + search.len()..]);
                        s_owned(out)
                    }
                }
            })
        }),
    );

    // ECMA-262 §22.1.3.19: replaceAll replaces every occurrence.
    string_fn(
        vm,
        "replaceAll",
        vec![str_t(), str_t(), str_t()],
        vec![str_t()],
        Box::new(|_ctx, args| {
            with_three_s_args(args, 0, 1, 2, |s, search, replace| {
                if !s.contains(search) {
                    s_val(s)
                } else {
                    s_owned(s.replace(search, replace))
                }
            })
        }),
    );
}

// ── Split ─────────────────────────────────────────────────────────

fn register_split(vm: &mut VM) {
    vm.register_host_fn(
        "ecma:string",
        "split",
        Box::new(|_ctx, args| {
            with_s_arg(args, 0, |s| {
                let separator = match args.get(1) {
                    Some(Value::String(text)) => Some(Cow::Borrowed(text.as_ref())),
                    None | Some(Value::Undefined) => None,
                    Some(other) => Some(crate::keys::value_display_cow(other)),
                };
                let limit: Option<usize> = match args.get(2) {
                    Some(Value::F64(n)) if *n >= 0.0 => Some(*n as usize),
                    Some(Value::I32(n)) if *n >= 0 => Some(*n as usize),
                    _ => None,
                };

                if matches!(limit, Some(0)) {
                    return Value::Object(vybe_runtime::heap::alloc(Object::new_array(Vec::new())));
                }
                let limit_count = limit.unwrap_or(usize::MAX);
                let parts: Vec<Value> = match separator {
                    None => vec![s_val(&s)],
                    Some(sep) if sep.is_empty() => {
                        // ECMA-262: empty string separator splits every char.
                        let char_count = s.chars().count().min(limit_count);
                        let mut parts = Vec::with_capacity(char_count);
                        for ch in s.chars().take(limit_count) {
                            parts.push(char_value(ch));
                        }
                        parts
                    }
                    Some(sep) => {
                        let mut parts = Vec::new();
                        if limit_count != usize::MAX {
                            parts.reserve(limit_count.min(16));
                        } else {
                            parts.reserve(4);
                        }
                        for piece in s.split(sep.as_ref()).take(limit_count) {
                            parts.push(s_val(piece));
                        }
                        parts
                    }
                };
                Value::Object(vybe_runtime::heap::alloc(Object::new_array(parts)))
            })
        }),
    );
}

// ── String constructor statics ────────────────────────────────────

fn register_constructor_statics(vm: &mut VM) {
    vm.register_host_fn(
        "ecma:string",
        "fromCharCode",
        Box::new(|_ctx, args| {
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
            s_owned(out)
        }),
    );

    vm.register_host_fn(
        "ecma:string",
        "fromCodePoint",
        Box::new(|_ctx, args| {
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
            s_owned(out)
        }),
    );
}

// ── localeCompare (ECMA-262 §22.1.3.10) ───────────────────────────
//
// `a.localeCompare(b)` — locale-aware comparison. Returns -1 / 0 / 1
// per ECMA-402; the spec allows any negative / zero / positive return.
// Without an Intl.Collator implementation, this falls back to the
// codepoint ordering Rust's `str::cmp` produces — matches Node's
// default behaviour when Intl isn't available.

fn register_locale_compare(vm: &mut VM) {
    // §22.1.3.12 allows `locales` and `options`; this handler implements
    // neither, so the declaration is the binary comparison it actually is.
    string_fn(
        vm,
        "localeCompare",
        vec![str_t(), str_t()],
        vec![ValType::I32],
        Box::new(|_ctx, args| {
            with_two_s_args(args, 0, 1, |a, b| {
                Value::I32(match a.cmp(b) {
                    std::cmp::Ordering::Less => -1,
                    std::cmp::Ordering::Equal => 0,
                    std::cmp::Ordering::Greater => 1,
                })
            })
        }),
    );
}

// ── normalize (ECMA-262 §22.1.3.13) ───────────────────────────────
//
// `s.normalize(form?)` — Unicode normalization. Form is one of NFC /
// NFD / NFKC / NFKD; defaults to NFC. We don't ship a Unicode
// normalization library, so MVP returns the input unchanged for ASCII
// input and signals lossless passthrough for the common case. Full
// implementation requires the `unicode-normalization` crate.

fn register_normalize(vm: &mut VM) {
    vm.register_host_fn(
        "ecma:string",
        "normalize",
        Box::new(|ctx, args| {
            with_s_arg(args, 0, |input| {
                let form = match args.get(1) {
                    None | Some(Value::Undefined) => "NFC",
                    Some(Value::String(form)) => form.as_ref(),
                    Some(other) => {
                        ctx.throw_value(crate::error::new_error(
                            ctx,
                            "RangeError",
                            &format!(
                                "The normalization form should be one of NFC, NFD, NFKC, NFKD: {}",
                                other
                            ),
                        ));
                        return Value::Null;
                    }
                };
                let normalized = match form {
                    "NFC" => input.nfc().collect::<String>(),
                    "NFD" => input.nfd().collect::<String>(),
                    "NFKC" => input.nfkc().collect::<String>(),
                    "NFKD" => input.nfkd().collect::<String>(),
                    _ => {
                        ctx.throw_value(crate::error::new_error(
                            ctx,
                            "RangeError",
                            &format!(
                                "The normalization form should be one of NFC, NFD, NFKC, NFKD: {}",
                                form
                            ),
                        ));
                        return Value::Null;
                    }
                };
                s_owned(normalized)
            })
        }),
    );
}

// ── ECMA-262 §19.2.6 URI globals ───────────────────────────────────
//
// `encodeURI`, `decodeURI`, `encodeURIComponent`, `decodeURIComponent`
// are spec'd as global functions but they're string transforms — they
// belong on `ecma:string` for the purposes of the host-fn registry.

fn decode_uri_string(input: &str) -> Result<String, &'static str> {
    let bytes = input.as_bytes();
    let mut out = Vec::with_capacity(bytes.len());
    let mut i = 0;
    while i < bytes.len() {
        if bytes[i] == b'%' {
            if i + 2 >= bytes.len() {
                return Err("URI malformed");
            }
            let hi = hex_nibble(bytes[i + 1]).ok_or("URI malformed")?;
            let lo = hex_nibble(bytes[i + 2]).ok_or("URI malformed")?;
            let byte = (hi << 4) | lo;
            out.push(byte);
            i += 3;
            continue;
        }
        out.push(bytes[i]);
        i += 1;
    }
    String::from_utf8(out).map_err(|_| "URI malformed")
}

#[inline]
fn uri_byte_is_unescaped(byte: u8, allow_reserved: bool) -> bool {
    byte.is_ascii_alphanumeric()
        || matches!(byte, b'-' | b'_' | b'.' | b'!' | b'~' | b'*' | b'\'' | b'(' | b')')
        || (allow_reserved
            && matches!(byte, b';' | b',' | b'/' | b'?' | b':' | b'@' | b'&' | b'=' | b'+' | b'$' | b'#'))
}

fn encode_uri_text(input: &str, allow_reserved: bool) -> Cow<'_, str> {
    let Some(first_escape) = input.bytes().position(|byte| !uri_byte_is_unescaped(byte, allow_reserved)) else {
        return Cow::Borrowed(input);
    };
    // Before the first escaped byte every byte is ASCII, so this is a UTF-8 boundary.
    let mut output = String::with_capacity(input.len());
    output.push_str(&input[..first_escape]);
    for &byte in &input.as_bytes()[first_escape..] {
        if uri_byte_is_unescaped(byte, allow_reserved) {
            output.push(byte as char);
        } else {
            push_percent_byte(&mut output, byte);
        }
    }
    Cow::Owned(output)
}

fn encode_uri_value(args: &[Value], allow_reserved: bool) -> Value {
    with_s_arg(args, 0, |input| match encode_uri_text(input, allow_reserved) {
        Cow::Borrowed(text) => match args.first() {
            Some(Value::String(original)) => Value::String(Arc::clone(original)),
            _ => s_val(text),
        },
        Cow::Owned(text) => s_owned(text),
    })
}

fn register_uri(vm: &mut VM) {
    // encodeURIComponent — encodes everything except the unreserved set
    // (ALPHA / DIGIT / `-` / `_` / `.` / `~` / `!` / `*` / `'` / `(` / `)`).
    // ECMA-262 §19.2.6.5.
    string_fn_free(
        vm,
        "encodeURIComponent",
        vec![str_t()],
        vec![str_t()],
        Box::new(|_ctx, args| encode_uri_value(args, false)),
    );

    // decodeURIComponent — reverses encodeURIComponent.
    string_fn_free(
        vm,
        "decodeURIComponent",
        vec![str_t()],
        vec![str_t()],
        Box::new(|ctx, args| {
            with_s_arg(args, 0, |s| match decode_uri_string(s) {
                Ok(decoded) => s_owned(decoded),
                Err(message) => {
                    ctx.throw_value(crate::error::new_error(ctx, "URIError", message));
                    Value::Undefined
                }
            })
        }),
    );

    // encodeURI — like encodeURIComponent but ALSO leaves URI-syntax
    // chars unencoded: `;` `,` `/` `?` `:` `@` `&` `=` `+` `$` `#`.
    // ECMA-262 §19.2.6.4.
    string_fn_free(
        vm,
        "encodeURI",
        vec![str_t()],
        vec![str_t()],
        Box::new(|_ctx, args| encode_uri_value(args, true)),
    );

    // decodeURI — reverses encodeURI. The spec defines a different
    // reserved set than decodeURIComponent (it preserves URI-syntax
    // chars even if they were percent-encoded), but for our MVP we
    // simply unescape every `%XX` — same behaviour as decodeURIComponent.
    string_fn_free(
        vm,
        "decodeURI",
        vec![str_t()],
        vec![str_t()],
        Box::new(|ctx, args| {
            with_s_arg(args, 0, |s| match decode_uri_string(s) {
                Ok(decoded) => s_owned(decoded),
                Err(message) => {
                    ctx.throw_value(crate::error::new_error(ctx, "URIError", message));
                    Value::Undefined
                }
            })
        }),
    );

    // Annex B `escape` — legacy percent encoder used by older JS code.
    // Leaves `A-Z a-z 0-9 @*_+-./` unescaped, encodes Latin-1 bytes as
    // `%XX`, and wider code points as `%uXXXX`.
    string_fn_free(
        vm,
        "escape",
        vec![str_t()],
        vec![str_t()],
        Box::new(|_ctx, args| {
            // ⛔ `escape`/`unescape` are GLOBAL functions, not methods, so
            // argument 0 is not a receiver they want — but under
            // `ReceiverAbi::Parameter` a plain call still puts one there
            // (§10.2.1.1 binds `undefined`). Reading the string at a fixed
            // index picked up that receiver and returned `undefined`. Inert
            // under the ambient binding.
            with_s_arg(args, 0, |s| {
                let mut encoded = String::with_capacity(s.len());
                for ch in s.chars() {
                    if ch.is_ascii_alphanumeric() || "@*_+-./".contains(ch) {
                        encoded.push(ch);
                    } else {
                        let code = ch as u32;
                        if code < 256 {
                            push_percent_byte(&mut encoded, code as u8);
                        } else {
                            push_percent_u16(&mut encoded, code);
                        }
                    }
                }
                s_owned(encoded)
            })
        }),
    );

    // Annex B `unescape` — reverses `%XX` and `%uXXXX` escapes.
    string_fn_free(
        vm,
        "unescape",
        vec![str_t()],
        vec![str_t()],
        Box::new(|_ctx, args| {
            // Global function, not a method — see `escape` above.
            with_s_arg(args, 0, |s| {
                let bytes = s.as_bytes();
                let mut out = String::with_capacity(bytes.len());
                let mut i = 0;
                while i < bytes.len() {
                    if bytes[i] == b'%' {
                        if i + 5 < bytes.len() && bytes[i + 1] == b'u' {
                            if let (Some(a), Some(b), Some(c), Some(d)) = (
                                hex_nibble(bytes[i + 2]),
                                hex_nibble(bytes[i + 3]),
                                hex_nibble(bytes[i + 4]),
                                hex_nibble(bytes[i + 5]),
                            ) {
                                let code = ((a as u32) << 12)
                                    | ((b as u32) << 8)
                                    | ((c as u32) << 4)
                                    | d as u32;
                                if let Some(ch) = char::from_u32(code) {
                                    out.push(ch);
                                    i += 6;
                                    continue;
                                }
                            }
                        } else if i + 2 < bytes.len() {
                            if let (Some(hi), Some(lo)) =
                                (hex_nibble(bytes[i + 1]), hex_nibble(bytes[i + 2]))
                            {
                                let byte = (hi << 4) | lo;
                                out.push(byte as char);
                                i += 3;
                                continue;
                            }
                        }
                    }
                    out.push(bytes[i] as char);
                    i += 1;
                }
                s_owned(out)
            })
        }),
    );
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
        if chunk.len() > 1 {
            result.push(BASE64_CHARS[((n >> 6) & 63) as usize] as char);
        } else {
            result.push('=');
        }
        if chunk.len() > 2 {
            result.push(BASE64_CHARS[(n & 63) as usize] as char);
        } else {
            result.push('=');
        }
    }
    result
}

fn base64_decode(s: &str) -> Option<Vec<u8>> {
    const DECODE: [u8; 128] = {
        let mut t = [255u8; 128];
        let mut i = 0u8;
        while i < 26 {
            t[(b'A' + i) as usize] = i;
            i += 1;
        }
        i = 0;
        while i < 26 {
            t[(b'a' + i) as usize] = 26 + i;
            i += 1;
        }
        i = 0;
        while i < 10 {
            t[(b'0' + i) as usize] = 52 + i;
            i += 1;
        }
        t[b'+' as usize] = 62;
        t[b'/' as usize] = 63;
        t
    };
    let filtered_len = s.bytes().filter(|b| !b.is_ascii_whitespace()).count();
    if filtered_len % 4 == 1 {
        return None;
    }
    let chunk_count = filtered_len / 4;
    let mut result = Vec::with_capacity(chunk_count * 3);
    let mut chunk = [0u8; 4];
    let mut chunk_len = 0usize;
    let mut chunk_index = 0usize;
    for byte in s.bytes().filter(|b| !b.is_ascii_whitespace()) {
        chunk[chunk_len] = byte;
        chunk_len += 1;
        if chunk_len != 4 {
            continue;
        }
        let last_chunk = chunk_index + 1 == chunk_count;
        if !base64_decode_chunk(&chunk, last_chunk, &DECODE, &mut result) {
            return None;
        }
        chunk_len = 0;
        chunk_index += 1;
    }
    if chunk_len != 0 {
        return None;
    }
    Some(result)
}

fn base64_decode_chunk(
    chunk: &[u8; 4],
    last_chunk: bool,
    decode: &[u8; 128],
    result: &mut Vec<u8>,
) -> bool {
    let a = chunk[0];
    let b = chunk[1];
    if a >= 128 || b >= 128 || a == b'=' || b == b'=' {
        return false;
    }
    let av = decode[a as usize] as u32;
    let bv = decode[b as usize] as u32;
    if av == 255 || bv == 255 {
        return false;
    }
    let c = chunk[2];
    let d = chunk[3];
    if c == b'=' {
        if d != b'=' || !last_chunk {
            return false;
        }
        result.push(((av << 2) | (bv >> 4)) as u8);
        return true;
    }
    if c >= 128 {
        return false;
    }
    let cv = decode[c as usize] as u32;
    if cv == 255 {
        return false;
    }
    result.push(((av << 2) | (bv >> 4)) as u8);
    result.push((((bv & 0xF) << 4) | (cv >> 2)) as u8);

    if d == b'=' {
        return last_chunk;
    }
    if d >= 128 {
        return false;
    }
    let dv = decode[d as usize] as u32;
    if dv == 255 {
        return false;
    }
    result.push((((cv & 0x3) << 6) | dv) as u8);
    true
}

fn latin1_string(bytes: &[u8]) -> String {
    if bytes.is_ascii() {
        return std::str::from_utf8(bytes).unwrap_or("").to_owned();
    }
    let mut out = String::with_capacity(bytes.len());
    for byte in bytes {
        out.push(char::from(*byte));
    }
    out
}

fn throw_type_error(ctx: &mut HostContext, message: &str) {
    ctx.throw_value(crate::error::new_error(ctx, "TypeError", message));
}

fn match_capture_values(regex: &regex::Regex, input: &str) -> Option<Vec<Value>> {
    // Preserve the cheap no-match path: captures() allocates capture storage even
    // when there is no match. Patterns without groups need no capture pass at all.
    let matched = regex.find(input)?;
    if regex.captures_len() == 1 {
        return Some(vec![crate::keys::string_value(matched.as_str())]);
    }
    let captures = regex.captures(input)?;
    let mut values = Vec::with_capacity(captures.len());
    for index in 0..captures.len() {
        values.push(match captures.get(index) {
            Some(capture) => crate::keys::string_value(capture.as_str()),
            None => Value::Undefined,
        });
    }
    Some(values)
}

fn register_base64(vm: &mut VM) {
    // btoa — base64-encode the input string's BYTES (treating it as
    // Latin-1 per HTML spec).
    string_fn(
        vm,
        "btoa",
        vec![str_t()],
        vec![str_t()],
        Box::new(|ctx, args| {
            with_s_arg(args, 0, |s| {
                if s.is_ascii() {
                    return s_owned(base64_encode(s.as_bytes()));
                }
                let mut bytes = Vec::with_capacity(s.len());
                for ch in s.chars() {
                    let code = ch as u32;
                    if code > 0xFF {
                        throw_type_error(
                            ctx,
                            "The string to be encoded contains characters outside of the Latin1 range",
                        );
                        return Value::Null;
                    }
                    bytes.push(code as u8);
                }
                s_owned(base64_encode(&bytes))
            })
        }),
    );

    // atob — base64-decode and interpret bytes as a Latin-1 string.
    string_fn(
        vm,
        "atob",
        vec![str_t()],
        vec![str_t()],
        Box::new(|ctx, args| {
            with_s_arg(args, 0, |s| match base64_decode(s) {
                Some(bytes) => s_owned(latin1_string(&bytes)),
                None => {
                    throw_type_error(ctx, "The string to be decoded is not correctly encoded");
                    Value::Null
                }
            })
        }),
    );

    // match(string, pattern) — §22.1.3.12. Returns first-match array or Null.
    string_fn(
        vm,
        "match",
        vec![str_t(), str_t()],
        vec![ValType::Any],
        Box::new(|_ctx, args| {
            with_two_s_args(args, 0, 1, |s, pattern| {
                if let Ok(re) = regex::Regex::new(pattern) {
                    if let Some(arr_vals) = match_capture_values(&re, s) {
                        return Value::Object(vybe_runtime::heap::alloc(Object::new_array(
                            arr_vals,
                        )));
                    }
                }
                Value::Null
            })
        }),
    );

    // search(string, pattern) — §22.1.3.21. Returns index of first match or -1.
    string_fn(
        vm,
        "search",
        vec![str_t(), str_t()],
        vec![ValType::F64],
        Box::new(|_ctx, args| {
            with_two_s_args(args, 0, 1, |s, pattern| {
                if let Ok(re) = regex::Regex::new(pattern) {
                    if let Some(m) = re.find(s) {
                        return Value::F64(m.start() as f64);
                    }
                }
                Value::F64(-1.0)
            })
        }),
    );

    // isWellFormed — §22.1.3.10 (ES2024). Rust strings are always UTF-8.
    // The handler ignores its operand — every string IS well-formed here — but
    // the receiver is still a parameter, and callers pass it.
    string_fn(
        vm,
        "isWellFormed",
        vec![str_t()],
        vec![ValType::Bool],
        Box::new(|_ctx, _args| Value::Bool(true)),
    );

    // toWellFormed — §22.1.3.31 (ES2024). Returns the string unchanged.
    string_fn(
        vm,
        "toWellFormed",
        vec![str_t()],
        vec![ValType::String],
        Box::new(|_ctx, args| args.first().cloned().unwrap_or(Value::Undefined)),
    );

    // raw(templateObject, ...subs) — §22.1.2.4.
    vm.register_host_fn(
        "ecma:string",
        "raw",
        Box::new(|_ctx, args| {
            // ECMA-262 §22.1.2.4: String.raw(template, ...subs)
            // Reads template.raw (the raw strings array), not the cooked template itself.
            let mut result = String::new();
            let subs = args.get(1..).unwrap_or(&[]);
            match args.first() {
                Some(Value::Object(obj)) => {
                    let o = obj.lock().unwrap();
                    // First try .raw property (tagged template path)
                    if let Some(raw_val) = o.properties.get("raw") {
                        if let Value::Object(raw_arr) = raw_val {
                            let raw = raw_arr.lock().unwrap();
                            if let ObjectKind::Array(v) = &raw.kind {
                                append_raw_template_parts(&mut result, v, subs);
                            }
                        }
                    } else {
                        // Fallback: use the array itself (cooked)
                        if let ObjectKind::Array(v) = &o.kind {
                            append_raw_template_parts(&mut result, v, subs);
                        }
                    }
                }
                _ => {}
            }
            s_val(&result)
        }),
    );

    // toLocaleString — same as toString for basic strings.
    string_fn(
        vm,
        "toLocaleString",
        vec![str_t()],
        vec![ValType::String],
        Box::new(|_ctx, args| args.first().cloned().unwrap_or(Value::Undefined)),
    );
}

/// §22.1.5.1 String.prototype[@@iterator]() — yields code points.
/// Strings are opaque in WASM (wasm:js-string spec); iteration
/// requires a host function, same as charCodeAt/codePointAt.
fn register_iterator(vm: &mut VM) {
    string_fn(
        vm,
        "iterator",
        vec![str_t()],
        vec![ValType::Any],
        Box::new(|_ctx, args| {
            let chars = with_s_arg(args, 0, |s| {
                let mut chars = Vec::with_capacity(s.len());
                for ch in s.chars() {
                    chars.push(char_value(ch));
                }
                chars
            });
            crate::array::make_array_iterator(chars)
        }),
    );
}

#[allow(dead_code)]
fn _force_object_use(_: ObjectKind) {}

#[cfg(test)]
mod uri_encoding_performance_tests {
    use super::*;

    fn old_encode(input: &str, reserved: bool) -> String {
        let mut output = String::with_capacity(input.len());
        let allowed = if reserved { "-_.!~*'();,/?:@&=+$#" } else { "-_.!~*'()" };
        for ch in input.chars() {
            if ch.is_ascii_alphanumeric() || allowed.contains(ch) {
                output.push(ch);
            } else {
                let mut bytes = [0u8; 4];
                for &byte in ch.encode_utf8(&mut bytes).as_bytes() {
                    push_percent_byte(&mut output, byte);
                }
            }
        }
        output
    }

    #[test]
    fn uri_encoder_preserves_ascii_and_unicode_results() {
        let ascii: String = (0..128).map(char::from).collect();
        for input in [ascii.as_str(), "", "safe-_.!~*'()", ";,/?:@&=+$#", "https://example.test/a?b=c#d", "\u{e9}\u{4e2d}\u{1f600}", "prefix \u{1f600}/suffix"] {
            for reserved in [false, true] {
                assert_eq!(encode_uri_text(input, reserved).as_ref(), old_encode(input, reserved));
            }
        }
        assert_eq!(encode_uri_text("\u{1f600}", false), "%F0%9F%98%80");
        assert_eq!(encode_uri_text("/", false), "%2F");
        assert_eq!(encode_uri_text("/", true), "/");
    }

    #[test]
    fn unchanged_uri_reuses_runtime_string() {
        let input = crate::keys::string_value("https://example.test/path?field=value#anchor");
        let output = encode_uri_value(std::slice::from_ref(&input), true);
        match (&input, &output) {
            (Value::String(a), Value::String(b)) => assert!(Arc::ptr_eq(a, b)),
            _ => panic!("string result expected"),
        }
        assert_eq!(encode_uri_value(&[Value::I32(123)], false), crate::keys::string_value("123"));
    }

    #[test]
    #[ignore = "matched URI adapter kernel benchmark; run explicitly with --nocapture"]
    fn uri_encoding_microbenchmark() {
        use std::time::Instant;
        let inputs = [
            "https://example.test/path?field=value#anchor",
            "component-with-only-unreserved-characters_123",
            "https://example.test/a path?q=\u{e9}\u{1f600}&x=1",
        ];
        for input in inputs {
            for reserved in [false, true] {
                let value = crate::keys::string_value(input);
                assert_eq!(encode_uri_value(std::slice::from_ref(&value), reserved), s_owned(old_encode(input, reserved)));
                for sample in 0..4 {
                    let mut old_time = std::time::Duration::ZERO;
                    let mut new_time = std::time::Duration::ZERO;
                    for optimized in if sample % 2 == 0 { [false, true] } else { [true, false] } {
                        let start = Instant::now();
                        for _ in 0..50_000 {
                            let result = if optimized {
                                encode_uri_value(std::slice::from_ref(&value), reserved)
                            } else {
                                s_owned(old_encode(input, reserved))
                            };
                            std::hint::black_box(result);
                        }
                        if optimized { new_time = start.elapsed(); } else { old_time = start.elapsed(); }
                    }
                    eprintln!("uri input={input:?} reserved={reserved} sample={sample} operations=50000 old={old_time:?} new={new_time:?}");
                }
            }
        }
    }
}

#[cfg(test)]
mod match_capture_performance_tests {
    use super::*;

    fn old_match(regex: &regex::Regex, input: &str) -> Option<Vec<Value>> {
        let matched = regex.find(input)?;
        let mut values = vec![crate::keys::string_value(matched.as_str())];
        if let Some(captures) = regex.captures(input) {
            for index in 1..captures.len() {
                values.push(match captures.get(index) {
                    Some(capture) => crate::keys::string_value(capture.as_str()),
                    None => Value::Undefined,
                });
            }
        }
        Some(values)
    }

    #[test]
    fn single_capture_search_preserves_existing_results() {
        for (pattern, input) in [
            ("(a)(b)?", "xa"),
            ("(a)(b)?", "xab"),
            ("a*", "bbb"),
            ("(?P<word>[a-z]+)-([0-9]+)", "item-42"),
            ("word", "not present"),
            ("\u{e9}+(\u{1f600})?", "x\u{e9}\u{e9}"),
            ("", ""),
        ] {
            let regex = regex::Regex::new(pattern).unwrap();
            assert_eq!(match_capture_values(&regex, input), old_match(&regex, input));
        }
        let regex = regex::Regex::new("(a)(b)?").unwrap();
        assert_eq!(match_capture_values(&regex, "a"), Some(vec![
            crate::keys::string_value("a"),
            crate::keys::string_value("a"),
            Value::Undefined,
        ]));
        assert!(regex::Regex::new("[").is_err());
    }

    #[test]
    #[ignore = "matched capture-search kernel benchmark; excludes regex compilation"]
    fn single_capture_search_microbenchmark() {
        use std::time::{Duration, Instant};
        let regex = regex::Regex::new("([a-z]+)-([0-9]+)$").unwrap();
        for matched in [false, true] {
            let mut input = "0123456789 ".repeat(4096);
            if matched { input.push_str("item-42"); }
            assert_eq!(match_capture_values(&regex, &input), old_match(&regex, &input));
            for sample in 0..4 {
                let mut old_time = Duration::ZERO;
                let mut new_time = Duration::ZERO;
                for optimized in if sample % 2 == 0 { [false, true] } else { [true, false] } {
                    let start = Instant::now();
                    for _ in 0..1000 {
                        std::hint::black_box(if optimized {
                            match_capture_values(&regex, &input)
                        } else {
                            old_match(&regex, &input)
                        });
                    }
                    if optimized { new_time = start.elapsed(); } else { old_time = start.elapsed(); }
                }
                eprintln!("captures matched={matched} sample={sample} operations=1000 old={old_time:?} new={new_time:?}");
            }
        }
    }
}

#[cfg(test)]
mod index_of_utf8_tests {
    use super::*;
    use vybe_runtime::{Chunk, Op};
    fn worker(old: bool) -> (VM, Value) {
        let mut owner = Chunk::new("index-owner");
        vybe_compiler::primitives::functions::emit_ref_func(&mut owner, 1, 0, 0);
        owner.emit_op(Op::RETURN, 0);
        let mut body = Chunk::new("index-search");
        body.arity = 3; body.param_count = 3; body.local_count = 3;
        for slot in 0..3 { body.emit_op_u16(Op::LOCAL_GET, slot, 0); }
        let import = body.add_import("ecma:string", "indexOf");
        body.emit_call(import, 3, 0); body.emit_op(Op::RETURN, 0);
        let mut chunks = vec![owner, body];
        vybe_compiler::primitives::bundle::finalize_with_runtime_helpers(&mut chunks);
        let mut vm = VM::new();
        vybe_compiler::primitives::platforms::register_platforms_all(&mut vm);
        if old {
    vm.register_host_fn(
        "ecma:string",
        "indexOf",
        Box::new(|_ctx, args| {
            // ECMA-262 §22.1.3.9: positions are UTF-16 code-unit indices,
            // consistent with `length` and `includes`. (Previously used
            // Rust `str::find`, which returns UTF-8 *byte* offsets — this
            // mis-indexed any non-ASCII string and could panic when `pos`
            // landed inside a multi-byte sequence.)
            with_two_s_args(args, 0, 1, |s, needle| {
                if s.is_ascii() && needle.is_ascii() {
                    let pos = i32_arg(args, 2, 0).max(0) as usize;
                    let start = pos.min(s.len());
                    if needle.is_empty() {
                        return Value::F64(start as f64);
                    }
                    return match s[start..].find(needle) {
                        Some(byte_idx) => Value::F64((start + byte_idx) as f64),
                        None => Value::F64(-1.0),
                    };
                }
                let hay = utf16_units(s);
                let ndl = utf16_units(needle);
                let pos = i32_arg(args, 2, 0).max(0) as usize;
                let start = pos.min(hay.len());
                if ndl.is_empty() {
                    return Value::F64(start as f64);
                }
                if ndl.len() > hay.len() {
                    return Value::F64(-1.0);
                }
                for i in start..=(hay.len() - ndl.len()) {
                    if hay[i..i + ndl.len()] == ndl[..] {
                        return Value::F64(i as f64);
                    }
                }
                Value::F64(-1.0)
            })
        }),
    );


        }
        let function = vm.run(chunks).unwrap(); (vm, function)
    }
    fn oracle(source: &str, needle: &str, position: usize) -> f64 {
        let hay: Vec<_> = source.encode_utf16().collect();
        let needle: Vec<_> = needle.encode_utf16().collect();
        let start = position.min(hay.len());
        if needle.is_empty() { return start as f64; }
        if needle.len() > hay.len() { return -1.0; }
        (start..=hay.len()-needle.len()).find(|&i| hay[i..i+needle.len()]==needle)
            .map(|i| i as f64).unwrap_or(-1.0)
    }
    #[test]
    fn allocation_free_index_of_preserves_actual_adapter_results() {
        let mut workers = [worker(true), worker(false)];
        for source in ["", "ABC", "Ωé中", "😀A😀", "𐐀Ω😀A", "A\0ΩA"] {
            for needle in ["", "A", "Ω", "😀", "𐐀", "\0", "Ω😀", "😀A", "missing"] {
                for position in [-1, 0, 1, 2, 3, 4, 5, 6, 20, i32::MAX] {
                    let args = [Value::String(Arc::from(source)), Value::String(Arc::from(needle)), Value::I32(position)];
                    let expected = Value::F64(oracle(source, needle, position.max(0) as usize));
                    for (vm, f) in &mut workers { assert_eq!(vm.invoke(f, &args).unwrap(), expected); assert!(vm.last_exception.is_none()); }
                }
            }
        }
        for receiver in [Value::Null, Value::Undefined, Value::I32(123), Value::Bool(true)] {
            for needle in [Value::Undefined, Value::Null, Value::I32(2), Value::Bool(false)] {
                for position in [Value::Undefined, Value::F64(f64::NAN), Value::F64(1.8), Value::F64(f64::INFINITY)] {
                    let args = [receiver.clone(), needle.clone(), position];
                    let before_fn = workers[0].1.clone();
                    let after_fn = workers[1].1.clone();
                    let before = workers[0].0.invoke(&before_fn, &args).unwrap();
                    let after = workers[1].0.invoke(&after_fn, &args).unwrap();
                    assert_eq!(before, after);
                }
            }
        }
    }
    #[test]
    #[ignore = "explicit opt0 adapter benchmark"]
    fn allocation_free_index_of_debug_pairs() {
        assert!(cfg!(debug_assertions));
        let mut workers = [worker(true), worker(false)];
        for (kind, source, needle, position, expected) in [
            ("ascii", "A".repeat(8192)+"!", "!", 0, 8192.0),
            ("unicode-scan", "Ω".repeat(8192)+"!", "!", 0, 8192.0),
            ("unicode-near", "Ω".repeat(8192), "Ω", 4096, 4096.0),
            ("astral-mid", "😀".repeat(4096), "😀", 4097, 4098.0)] {
            let args = [Value::String(Arc::from(source)), Value::String(Arc::from(needle)), Value::I32(position)];
            let batch = |worker: &mut (VM, Value), count: usize| {
                let start = std::time::Instant::now();
                for _ in 0..count { assert_eq!(worker.0.invoke(&worker.1, &args).unwrap(), Value::F64(expected)); }
                start.elapsed().as_nanos()
            };
            let mut counts = [1usize; 2];
            for side in 0..2 { while batch(&mut workers[side], counts[side]) < 500_000_000 { counts[side] *= 2; } }
            let mut samples = Vec::new(); let mut ratios = Vec::new();
            for pair in 0..7 {
                let mut ns = [0u128; 2];
                for side in if pair%2==0 {[0,1]} else {[1,0]} {
                    ns[side] = batch(&mut workers[side], counts[side]); samples.push((pair,side,counts[side],ns[side]));
                }
                ratios.push((ns[0] as f64/counts[0] as f64)/(ns[1] as f64/counts[1] as f64));
            }
            let mut sorted = ratios.clone(); sorted.sort_by(f64::total_cmp);
            eprintln!("indexOf {kind}: {:.4}x; ratios={ratios:?}; raw={samples:?}",sorted[3]);
        }
    }
}

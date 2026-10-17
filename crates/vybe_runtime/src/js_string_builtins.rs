//! # `wasm:js-string`
//!
//! **Merged** js-string-builtins WebAssembly proposal (V8 implements natively).
//! Spec: `proposals/js-string-builtins/proposals/js-string-builtins/Overview.md`
//!
//! Also includes the `wasm:js-string` extensions from the Stage-1
//! js-primitive-builtins proposal: `fromI32`, `fromU32`, `fromI64`, `fromU64`,
//! `fromF64`.
//!
//! All indices and lengths are in UTF-16 code units, matching JS semantics.
//! Functions marked `trap()` in the spec call `ctx.throw_value(...)` here.

use crate::value::ObjectKind;
use crate::{HostContext, VM, Value};
use std::sync::{Arc, Mutex};

/// Preserve the original single-string buffer path, retaining two older views.
/// Only the current string is strongly retained; older identities are weak.
#[derive(Default)]
struct Utf16Entry {
    source: Option<std::sync::Weak<str>>,
    units: Vec<u16>,
    ascii: bool,
}
#[derive(Default)]
struct Utf16View {
    source: Option<Arc<str>>,
    units: Vec<u16>,
    ascii: bool,
    previous: [Utf16Entry; 2],
    oldest: usize,
}
impl Utf16View {
    fn select(&mut self, source: &Arc<str>) {
        if self.source.as_ref().is_some_and(|old| Arc::ptr_eq(old, source)) {
            return;
        }
        let found = self.previous.iter().position(|entry| {
            entry.source.as_ref().is_some_and(|old| {
                std::ptr::eq(old.as_ptr(), Arc::as_ptr(source))
            })
        });
        let slot = found.unwrap_or(self.oldest);
        let entry = &mut self.previous[slot];
        std::mem::swap(&mut self.units, &mut entry.units);
        std::mem::swap(&mut self.ascii, &mut entry.ascii);
        entry.source = self.source.as_ref().map(Arc::downgrade);
        if found.is_none() {
            self.ascii = source.is_ascii();
            self.units.clear();
            if !self.ascii {
                self.units.extend(source.encode_utf16());
            }
        }
        self.oldest = 1 - slot;
        self.source = Some(source.clone());
        // A dead identity can never be queried again. Release its weak
        // allocation and decoded capacity when another string is selected.
        for entry in &mut self.previous {
            if entry.source.as_ref().is_some_and(|old| old.strong_count() == 0) {
                entry.source = None;
                entry.units = Vec::new();
                entry.ascii = false;
            }
        }
    }

    #[cfg(test)]
    fn units(&mut self, source: &Arc<str>) -> &[u16] {
        self.select(source);
        if self.ascii && self.units.len() != source.len() {
            self.units.extend(source.bytes().map(u16::from));
        }
        &self.units
    }
}

fn trap(ctx: &mut HostContext, msg: &str) {
    ctx.throw_value(Value::String(Arc::from(msg)));
}

fn is_string(v: &Value) -> bool {
    matches!(v, Value::String(_))
}

// Snapshot only the requested range. Conversion stays outside the object lock:
// dynamic compatibility values may themselves refer to locked objects.
// Keep Vec's original full-range clone path as well as the selected-slice path.
// Accepting Vec here is intentional: this helper owns both clone strategies.
#[inline]
fn snapshot_char_code_range(elems: &Vec<Value>, start: usize, end: usize) -> Option<Vec<Value>> {
    if start == 0 && end == elems.len() {
        Some(elems.clone())
    } else {
        elems.get(start..end).map(<[Value]>::to_vec)
    }
}

fn copy_char_code_units(source: &str, elems: &mut [Value], start: usize) -> Option<usize> {
    let count = crate::string_lengths::utf16_units(source);
    let end = start.checked_add(count)?;
    let destination = elems.get_mut(start..end)?;
    // For valid UTF-8, every non-ASCII scalar occupies more bytes than
    // UTF-16 units. Reuse the count instead of scanning for ASCII again.
    if count == source.len() {
        for (slot, byte) in destination.iter_mut().zip(source.bytes()) {
            *slot = Value::I32(byte as i32);
        }
    } else {
        for (slot, unit) in destination.iter_mut().zip(source.encode_utf16()) {
            *slot = Value::I32(unit as i32);
        }
    }
    Some(count)
}

pub fn register(vm: &mut VM) {
    // js-string-builtins UTF-8 extension. Rust strings already store UTF-8.
    vm.register_host_fn(
        "wasm:text-encoder",
        "measureStringAsUTF8",
        Box::new(|ctx: &mut HostContext, args: &[Value]| match args.first() {
            Some(Value::String(s)) => match u32::try_from(s.len()) {
                Ok(length) => Value::I32(length as i32),
                Err(_) => {
                    trap(ctx, "wasm:text-encoder.measureStringAsUTF8: length exceeds u32");
                    Value::Null
                }
            },
            _ => {
                trap(ctx, "wasm:text-encoder.measureStringAsUTF8: not a string");
                Value::Null
            }
        }),
    );
    let indexed_view = Arc::new(Mutex::new(Utf16View::default()));
    // test(externref) -> i32
    // Returns 1 if string, 0 otherwise (null also returns 0 per spec).
    vm.register_host_fn(
        "wasm:js-string",
        "test",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            Value::I32(if matches!(args.first(), Some(Value::String(_))) {
                1
            } else {
                0
            })
        }),
    );

    // cast(externref) -> (ref extern)
    // Traps on null or non-string.
    vm.register_host_fn(
        "wasm:js-string",
        "cast",
        Box::new(|ctx: &mut HostContext, args: &[Value]| match args.first() {
            Some(v @ Value::String(_)) => v.clone(),
            _ => {
                trap(ctx, "TypeError: wasm:js-string.cast — not a string");
                Value::Null
            }
        }),
    );

    // length(externref) -> i32
    // UTF-16 code unit count. Traps on null or non-string.
    vm.register_host_fn(
        "wasm:js-string",
        "length",
        Box::new(|ctx: &mut HostContext, args: &[Value]| match args.first() {
            Some(Value::String(s)) => Value::I32(crate::string_lengths::utf16_units(s) as i32),
            _ => {
                trap(ctx, "TypeError: wasm:js-string.length — not a string");
                Value::Null
            }
        }),
    );

    // concat(externref, externref) -> (ref extern)
    // Traps if either argument is null or non-string.
    vm.register_host_fn(
        "wasm:js-string",
        "concat",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            let a = match args.first() {
                Some(Value::String(s)) => s,
                _ => {
                    trap(
                        ctx,
                        "TypeError: wasm:js-string.concat — first arg not a string",
                    );
                    return Value::Null;
                }
            };
            let b = match args.get(1) {
                Some(Value::String(s)) => s,
                _ => {
                    trap(
                        ctx,
                        "TypeError: wasm:js-string.concat — second arg not a string",
                    );
                    return Value::Null;
                }
            };
            if a.is_empty() {
                return Value::String(b.clone());
            }
            if b.is_empty() {
                return Value::String(a.clone());
            }
            Value::String(crate::string_concat::concat2_arc(a, b))
        }),
    );

    // substring(externref, i32, i32) -> (ref extern)
    // start and end are treated as u32 (>>>= 0), clamped to [0, length].
    // If start > end they are swapped (JS String.prototype.substring semantics).
    // Operates on UTF-16 code units. Traps on null or non-string.
    let substring_view = indexed_view.clone();
    vm.register_host_fn(
        "wasm:js-string",
        "substring",
        Box::new(move |ctx: &mut HostContext, args: &[Value]| {
            let s = match args.first() {
                Some(Value::String(s)) => s.clone(),
                _ => {
                    trap(ctx, "TypeError: wasm:js-string.substring — not a string");
                    return Value::Null;
                }
            };
            let mut view = substring_view.lock().unwrap();
            view.select(&s);
            let len = if view.ascii { s.len() } else { view.units.len() };
            let start = (args.get(1).map(|v| v.as_i32()).unwrap_or(0) as u32) as usize;
            let end = (args.get(2).map(|v| v.as_i32()).unwrap_or(0) as u32) as usize;
            let start = start.min(len);
            let end = end.min(len);
            let (start, end) = if start > end {
                (end, start)
            } else {
                (start, end)
            };
            if view.ascii {
                // ASCII byte positions equal UTF-16 code-unit positions.
                return Value::String(Arc::from(&s[start..end]));
            }
            Value::String(Arc::from(
                String::from_utf16_lossy(&view.units[start..end]).as_ref(),
            ))
        }),
    );

    // equals(externref, externref) -> i32
    // Allows null (null == null is 1, null != string is 0).
    // Traps if non-null and non-string.
    vm.register_host_fn(
        "wasm:js-string",
        "equals",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            let a = args.first().unwrap_or(&Value::Null);
            let b = args.get(1).unwrap_or(&Value::Null);
            if !matches!(a, Value::Null) && !is_string(a) {
                trap(
                    ctx,
                    "TypeError: wasm:js-string.equals — first arg not a string or null",
                );
                return Value::Null;
            }
            if !matches!(b, Value::Null) && !is_string(b) {
                trap(
                    ctx,
                    "TypeError: wasm:js-string.equals — second arg not a string or null",
                );
                return Value::Null;
            }
            Value::I32(match (a, b) {
                (Value::String(x), Value::String(y)) => {
                    if x == y {
                        1
                    } else {
                        0
                    }
                }
                (Value::Null, Value::Null) => 1,
                _ => 0,
            })
        }),
    );

    // compare(externref, externref) -> i32 (-1, 0, 1)
    // Traps on null or non-string (no meaningful ordering for null per spec).
    vm.register_host_fn(
        "wasm:js-string",
        "compare",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            let a = match args.first() {
                Some(Value::String(s)) => s.clone(),
                _ => {
                    trap(
                        ctx,
                        "TypeError: wasm:js-string.compare — first arg not a string",
                    );
                    return Value::Null;
                }
            };
            let b = match args.get(1) {
                Some(Value::String(s)) => s.clone(),
                _ => {
                    trap(
                        ctx,
                        "TypeError: wasm:js-string.compare — second arg not a string",
                    );
                    return Value::Null;
                }
            };
            Value::I32(a.cmp(&b) as i32)
        }),
    );

    // charCodeAt(externref, i32) -> i32
    // index treated as u32. Returns the UTF-16 code unit at that position.
    // Traps on null/non-string, traps if index >= length.
    let char_view = indexed_view.clone();
    vm.register_host_fn(
        "wasm:js-string",
        "charCodeAt",
        Box::new(move |ctx: &mut HostContext, args: &[Value]| {
            let s = match args.first() {
                Some(Value::String(s)) => s.clone(),
                _ => {
                    trap(ctx, "TypeError: wasm:js-string.charCodeAt — not a string");
                    return Value::Null;
                }
            };
            let mut view = char_view.lock().unwrap();
            view.select(&s);
            let idx = (args.get(1).map(|v| v.as_i32()).unwrap_or(0) as u32) as usize;
            let len = if view.ascii { s.len() } else { view.units.len() };
            if idx >= len {
                trap(
                    ctx,
                    "RangeError: wasm:js-string.charCodeAt — index out of bounds",
                );
                return Value::Null;
            }
            Value::I32(if view.ascii { s.as_bytes()[idx] as i32 } else { view.units[idx] as i32 })
        }),
    );

    // codePointAt(externref, i32) -> i32
    // index treated as u32. Returns the full Unicode code point starting at that
    // UTF-16 position (handles surrogate pairs). Traps on null/non-string or OOB.
    vm.register_host_fn(
        "wasm:js-string",
        "codePointAt",
        Box::new(move |ctx: &mut HostContext, args: &[Value]| {
            let s = match args.first() {
                Some(Value::String(s)) => s.clone(),
                _ => {
                    trap(ctx, "TypeError: wasm:js-string.codePointAt — not a string");
                    return Value::Null;
                }
            };
            let mut view = indexed_view.lock().unwrap();
            view.select(&s);
            let idx = (args.get(1).map(|v| v.as_i32()).unwrap_or(0) as u32) as usize;
            let len = if view.ascii { s.len() } else { view.units.len() };
            if idx >= len {
                trap(
                    ctx,
                    "RangeError: wasm:js-string.codePointAt — index out of bounds",
                );
                return Value::Null;
            }
            if view.ascii {
                return Value::I32(s.as_bytes()[idx] as i32);
            }
            let units = &view.units;
            let hi = units[idx];
            if (0xD800..0xDC00).contains(&hi) {
                if let Some(&lo) = units.get(idx + 1) {
                    if (0xDC00..=0xDFFF).contains(&lo) {
                        let cp = 0x10000u32 + ((hi as u32 - 0xD800) << 10) + (lo as u32 - 0xDC00);
                        return Value::I32(cp as i32);
                    }
                }
            }
            Value::I32(hi as i32)
        }),
    );

    // fromCharCode(i32) -> (ref extern)
    // charCode treated as u32 (>>>= 0). Returns single-char string.
    vm.register_host_fn(
        "wasm:js-string",
        "fromCharCode",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let code = args.first().map(|v| v.as_i32()).unwrap_or(0) as u32 as u16;
            let s = String::from_utf16_lossy(&[code]);
            Value::String(Arc::from(s.as_ref()))
        }),
    );

    // fromCodePoint(i32) -> (ref extern)
    // codePoint treated as u32. Traps if > 0x10FFFF.
    vm.register_host_fn(
        "wasm:js-string",
        "fromCodePoint",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            let code = args.first().map(|v| v.as_i32()).unwrap_or(0) as u32;
            if code > 0x10FFFF {
                trap(
                    ctx,
                    "RangeError: wasm:js-string.fromCodePoint — code point out of range",
                );
                return Value::Null;
            }
            match char::from_u32(code) {
                Some(c) => Value::String(Arc::from(c.to_string().as_str())),
                None => {
                    trap(
                        ctx,
                        "RangeError: wasm:js-string.fromCodePoint — invalid code point",
                    );
                    Value::Null
                }
            }
        }),
    );

    // fromCharCodeArray(array, i32, i32) -> (ref extern)
    // Converts range [start, end) of a mutable i16 array to a string,
    // treating each element as an unsigned 16-bit char code.
    // Traps if array is null, range is invalid, or range is out of bounds.
    vm.register_host_fn(
        "wasm:js-string",
        "fromCharCodeArray",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            let arr = match args.first() {
                Some(Value::Object(o)) => o.clone(),
                _ => {
                    trap(
                        ctx,
                        "TypeError: wasm:js-string.fromCharCodeArray — not an array",
                    );
                    return Value::Null;
                }
            };
            let obj = arr.lock().unwrap();
            let elems = match &obj.kind {
                ObjectKind::Array(e) => e,
                _ => {
                    drop(obj);
                    trap(
                        ctx,
                        "TypeError: wasm:js-string.fromCharCodeArray — not an array",
                    );
                    return Value::Null;
                }
            };
            let start = (args.get(1).map(|v| v.as_i32()).unwrap_or(0) as u32) as usize;
            let end = (args.get(2).map(|v| v.as_i32()).unwrap_or(0) as u32) as usize;
            let Some(elems) = snapshot_char_code_range(elems, start, end) else {
                drop(obj);
                trap(
                    ctx,
                    "RangeError: wasm:js-string.fromCharCodeArray — range out of bounds",
                );
                return Value::Null;
            };
            drop(obj);
            let units: Vec<u16> = elems
                .iter()
                .map(|v| v.as_i32() as u16)
                .collect();
            Value::String(Arc::from(String::from_utf16_lossy(&units).as_ref()))
        }),
    );

    // intoCharCodeArray(externref, array, i32) -> i32
    // Copies the UTF-16 code units of string into array starting at start.
    // Returns number of code units written. Traps if string doesn't fit.
    vm.register_host_fn(
        "wasm:js-string",
        "intoCharCodeArray",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            let s = match args.first() {
                Some(Value::String(s)) => s.clone(),
                _ => {
                    trap(
                        ctx,
                        "TypeError: wasm:js-string.intoCharCodeArray — not a string",
                    );
                    return Value::Null;
                }
            };
            let arr = match args.get(1) {
                Some(Value::Object(o)) => o.clone(),
                _ => {
                    trap(
                        ctx,
                        "TypeError: wasm:js-string.intoCharCodeArray — not an array",
                    );
                    return Value::Null;
                }
            };
            let start = (args.get(2).map(|v| v.as_i32()).unwrap_or(0) as u32) as usize;
            let mut obj = arr.lock().unwrap();
            let elems = match &mut obj.kind {
                ObjectKind::Array(e) => e,
                _ => {
                    drop(obj);
                    trap(
                        ctx,
                        "TypeError: wasm:js-string.intoCharCodeArray — not an array",
                    );
                    return Value::Null;
                }
            };
            let Some(count) = copy_char_code_units(&s, elems, start) else {
                drop(obj);
                trap(
                    ctx,
                    "RangeError: wasm:js-string.intoCharCodeArray — string doesn't fit in array",
                );
                return Value::Null;
            };
            Value::I32(count as i32)
        }),
    );

    // ── Extensions from js-primitive-builtins (wasm:js-string additions) ──

    // fromI32(i32) -> (ref extern) — signed decimal string
    vm.register_host_fn(
        "wasm:js-string",
        "fromI32",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            Value::String(Arc::from(format!(
                "{}",
                args.first().map(|v| v.as_i32()).unwrap_or(0)
            )))
        }),
    );

    // fromU32(i32) -> (ref extern) — unsigned decimal string (reinterpret i32 as u32)
    vm.register_host_fn(
        "wasm:js-string",
        "fromU32",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let n = args.first().map(|v| v.as_i32()).unwrap_or(0) as u32;
            Value::String(Arc::from(format!("{}", n)))
        }),
    );

    // fromI64(i64) -> (ref extern) — signed decimal string
    vm.register_host_fn(
        "wasm:js-string",
        "fromI64",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            Value::String(Arc::from(format!(
                "{}",
                args.first().map(|v| v.as_i64()).unwrap_or(0)
            )))
        }),
    );

    // fromU64(i64) -> (ref extern) — unsigned decimal string (reinterpret i64 as u64)
    vm.register_host_fn(
        "wasm:js-string",
        "fromU64",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let n = args.first().map(|v| v.as_i64()).unwrap_or(0) as u64;
            Value::String(Arc::from(format!("{}", n)))
        }),
    );

    // fromF64(f64) -> (ref extern) — JS-style number-to-string
    vm.register_host_fn(
        "wasm:js-string",
        "fromF64",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let n = args.first().map(|v| v.as_f64()).unwrap_or(0.0);
            let s = if n.is_nan() {
                "NaN".to_string()
            } else if n.is_infinite() {
                if n > 0.0 {
                    "Infinity".to_string()
                } else {
                    "-Infinity".to_string()
                }
            } else {
                format!("{}", n)
            };
            Value::String(Arc::from(s))
        }),
    );
}

#[cfg(test)]
#[path = "js_string_ascii_view_tests.rs"]
mod ascii_view_tests;

#[cfg(test)]
#[path = "js_string_range_tests.rs"]
mod range_tests;

#[cfg(test)]
mod utf16_view_tests {
use super::Utf16View;
use std::sync::Arc;
#[test]
fn every_unicode_scalar_and_interleaved_tables_keep_exact_utf16() {
 let mut view=Utf16View::default();let table:Arc<str>=Arc::from("AZéΩ😀\0");let second:Arc<str>=Arc::from("ß中𐐀");
 for cp in 0..=0x10ffff {if let Some(c)=char::from_u32(cp) {let s:Arc<str>=Arc::from(c.to_string());let expected:Vec<_>=s.encode_utf16().collect();assert_eq!(view.units(&s),expected);assert_eq!(view.units(&table),table.encode_utf16().collect::<Vec<_>>());assert_eq!(view.units(&second),second.encode_utf16().collect::<Vec<_>>());assert_eq!(view.units(&s),expected);}}
}
#[test]
fn cached_old_views_do_not_retain_string_roots() {
 let mut view=Utf16View::default();let a:Arc<str>=Arc::from("first");let weak=Arc::downgrade(&a);view.units(&a);drop(a);assert!(weak.upgrade().is_some());let b:Arc<str>=Arc::from("second");view.units(&b);assert!(weak.upgrade().is_none());let w=Arc::downgrade(&b);drop(b);assert!(w.upgrade().is_some());drop(view);assert!(w.upgrade().is_none());
}
#[test]
fn three_live_strings_reuse_their_buffers_without_retranscoding() {
 let texts:[Arc<str>;3]=[Arc::from("é".repeat(4096)),Arc::from("Ω".repeat(128)),Arc::from("😀".repeat(128))];let mut view=Utf16View::default();let ptrs:Vec<_>=texts.iter().map(|s|view.units(s).as_ptr()).collect();for _ in 0..100 {for (s,ptr) in texts.iter().zip(&ptrs) {assert_eq!(view.units(s).as_ptr(),*ptr);}}
}
#[test]
fn expired_identity_and_eviction_do_not_alias_equal_or_new_strings() {
 let mut view=Utf16View::default();for round in 0..10000 {let s:Arc<str>=Arc::from(format!("{round}😀"));let expected:Vec<_>=s.encode_utf16().collect();assert_eq!(view.units(&s),expected);let equal:Arc<str>=Arc::from(s.to_string());assert!(!Arc::ptr_eq(&s,&equal));assert_eq!(view.units(&equal),expected);let third:Arc<str>=Arc::from("x");view.units(&third);let fourth:Arc<str>=Arc::from("\0é");view.units(&fourth);assert_eq!(view.units(&s),expected);}
}

}

#[cfg(test)]
mod utf16_expired_view_tests {
use super::Utf16View;
use std::sync::Arc;
#[test] fn expired_secondary_capacity_is_released_after_switch(){let mut v=Utf16View::default();let big:Arc<str>=Arc::from("Ω".repeat(1048576));let weak=Arc::downgrade(&big);v.units(&big);drop(big);let small:Arc<str>=Arc::from("x");assert_eq!(v.units(&small),[120]);assert!(weak.upgrade().is_none());assert!(v.previous.iter().all(|e|e.source.is_none()||e.source.as_ref().unwrap().strong_count()>0));assert!(v.previous.iter().filter(|e|e.source.is_none()).all(|e|e.units.capacity()==0));}
#[test] fn live_three_string_buffers_and_roots_are_preserved(){let texts:[Arc<str>;3]=[Arc::from("Ω".repeat(4096)),Arc::from("AZé"),Arc::from("ß𐐀")];let mut v=Utf16View::default();let ptrs:Vec<_>=texts.iter().map(|s|v.units(s).as_ptr()).collect();for _ in 0..100{for (s,p) in texts.iter().zip(&ptrs){assert_eq!(v.units(s).as_ptr(),*p);}}}

}

#[cfg(all(test, feature = "primitive-benchmarks"))]
#[path = "js_string_length_cache_tests.rs"]
mod length_cache_tests;

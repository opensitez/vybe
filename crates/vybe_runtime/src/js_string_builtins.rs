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

use std::sync::Arc;
use crate::value::ObjectKind;
use crate::{HostContext, VM, Value};

fn trap(ctx: &mut HostContext, msg: &str) {
    ctx.throw_value(Value::String(Arc::from(msg)));
}

fn is_string(v: &Value) -> bool {
    matches!(v, Value::String(_))
}

pub fn register(vm: &mut VM) {
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
            Some(Value::String(s)) => Value::I32(s.encode_utf16().count() as i32),
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
                Some(Value::String(s)) => s.clone(),
                _ => {
                    trap(
                        ctx,
                        "TypeError: wasm:js-string.concat — first arg not a string",
                    );
                    return Value::Null;
                }
            };
            let b = match args.get(1) {
                Some(Value::String(s)) => s.clone(),
                _ => {
                    trap(
                        ctx,
                        "TypeError: wasm:js-string.concat — second arg not a string",
                    );
                    return Value::Null;
                }
            };
            Value::String(Arc::from(format!("{}{}", a, b).as_str()))
        }),
    );

    // substring(externref, i32, i32) -> (ref extern)
    // start and end are treated as u32 (>>>= 0), clamped to [0, length].
    // If start > end they are swapped (JS String.prototype.substring semantics).
    // Operates on UTF-16 code units. Traps on null or non-string.
    vm.register_host_fn(
        "wasm:js-string",
        "substring",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            let s = match args.first() {
                Some(Value::String(s)) => s.clone(),
                _ => {
                    trap(ctx, "TypeError: wasm:js-string.substring — not a string");
                    return Value::Null;
                }
            };
            let units: Vec<u16> = s.encode_utf16().collect();
            let len = units.len();
            let start = (args.get(1).map(|v| v.as_i32()).unwrap_or(0) as u32) as usize;
            let end = (args.get(2).map(|v| v.as_i32()).unwrap_or(0) as u32) as usize;
            let start = start.min(len);
            let end = end.min(len);
            let (start, end) = if start > end {
                (end, start)
            } else {
                (start, end)
            };
            Value::String(Arc::from(
                String::from_utf16_lossy(&units[start..end]).as_ref(),
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
                _ => 0 })
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
    vm.register_host_fn(
        "wasm:js-string",
        "charCodeAt",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            let s = match args.first() {
                Some(Value::String(s)) => s.clone(),
                _ => {
                    trap(ctx, "TypeError: wasm:js-string.charCodeAt — not a string");
                    return Value::Null;
                }
            };
            let units: Vec<u16> = s.encode_utf16().collect();
            let idx = (args.get(1).map(|v| v.as_i32()).unwrap_or(0) as u32) as usize;
            if idx >= units.len() {
                trap(
                    ctx,
                    "RangeError: wasm:js-string.charCodeAt — index out of bounds",
                );
                return Value::Null;
            }
            Value::I32(units[idx] as i32)
        }),
    );

    // codePointAt(externref, i32) -> i32
    // index treated as u32. Returns the full Unicode code point starting at that
    // UTF-16 position (handles surrogate pairs). Traps on null/non-string or OOB.
    vm.register_host_fn(
        "wasm:js-string",
        "codePointAt",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            let s = match args.first() {
                Some(Value::String(s)) => s.clone(),
                _ => {
                    trap(ctx, "TypeError: wasm:js-string.codePointAt — not a string");
                    return Value::Null;
                }
            };
            let units: Vec<u16> = s.encode_utf16().collect();
            let idx = (args.get(1).map(|v| v.as_i32()).unwrap_or(0) as u32) as usize;
            if idx >= units.len() {
                trap(
                    ctx,
                    "RangeError: wasm:js-string.codePointAt — index out of bounds",
                );
                return Value::Null;
            }
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
                ObjectKind::Array(e) => e.clone(),
                _ => {
                    drop(obj);
                    trap(
                        ctx,
                        "TypeError: wasm:js-string.fromCharCodeArray — not an array",
                    );
                    return Value::Null;
                }
            };
            drop(obj);
            let arr_len = elems.len();
            let start = (args.get(1).map(|v| v.as_i32()).unwrap_or(0) as u32) as usize;
            let end = (args.get(2).map(|v| v.as_i32()).unwrap_or(0) as u32) as usize;
            if start > end || end > arr_len {
                trap(
                    ctx,
                    "RangeError: wasm:js-string.fromCharCodeArray — range out of bounds",
                );
                return Value::Null;
            }
            let units: Vec<u16> = elems[start..end]
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
            let units: Vec<u16> = s.encode_utf16().collect();
            let count = units.len();
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
            if start + count > elems.len() {
                drop(obj);
                trap(
                    ctx,
                    "RangeError: wasm:js-string.intoCharCodeArray — string doesn't fit in array",
                );
                return Value::Null;
            }
            for (i, &unit) in units.iter().enumerate() {
                elems[start + i] = Value::I32(unit as i32);
            }
            Value::I32(count as i32)
        }),
    );

    // ── Extensions from js-primitive-builtins (wasm:js-string additions) ──

    // fromI32(i32) -> (ref extern) — signed decimal string
    vm.register_host_fn(
        "wasm:js-string",
        "fromI32",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            Value::String(Arc::from(
                format!("{}", args.first().map(|v| v.as_i32()).unwrap_or(0)).as_str(),
            ))
        }),
    );

    // fromU32(i32) -> (ref extern) — unsigned decimal string (reinterpret i32 as u32)
    vm.register_host_fn(
        "wasm:js-string",
        "fromU32",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let n = args.first().map(|v| v.as_i32()).unwrap_or(0) as u32;
            Value::String(Arc::from(format!("{}", n).as_str()))
        }),
    );

    // fromI64(i64) -> (ref extern) — signed decimal string
    vm.register_host_fn(
        "wasm:js-string",
        "fromI64",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            Value::String(Arc::from(
                format!("{}", args.first().map(|v| v.as_i64()).unwrap_or(0)).as_str(),
            ))
        }),
    );

    // fromU64(i64) -> (ref extern) — unsigned decimal string (reinterpret i64 as u64)
    vm.register_host_fn(
        "wasm:js-string",
        "fromU64",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let n = args.first().map(|v| v.as_i64()).unwrap_or(0) as u64;
            Value::String(Arc::from(format!("{}", n).as_str()))
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
            Value::String(Arc::from(s.as_str()))
        }),
    );
}

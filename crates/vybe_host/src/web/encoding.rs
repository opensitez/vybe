//! WHATWG Encoding Living Standard — TextEncoder + TextDecoder.
//!
//!   `new TextEncoder()` — encodes JS strings (UTF-16) to UTF-8 bytes.
//!   `encoder.encode(input)` → Uint8Array
//!   `encoder.encodeInto(input, dest)` → { read, written }
//!
//!   `new TextDecoder(label?, options?)` — decodes bytes to a JS string.
//!   `decoder.decode(input?, options?)` → string
//!
//! Vybe stores strings as Rust `Arc<str>` (UTF-8 internally), so the
//! encode path is essentially the identity over `as_bytes()`. Decode
//! must validate UTF-8; invalid sequences become U+FFFD by default
//! per spec §10.1, or throw if `fatal: true`.

use std::str;
use std::sync::{Arc, Mutex};
use vybe_bytecode::value::{Object, ObjectKind, TypedElemKind};
use vybe_bytecode::{HostContext, VM, Value};

fn make_uint8_array(bytes: Vec<u8>) -> Value {
    let array = crate::ecma::typedarray::new_typed_array(TypedElemKind::U8, bytes.len());
    if let Value::Object(obj) = &array {
        obj.lock()
            .unwrap()
            .properties
            .insert("__type".into(), Value::String(Arc::from("Uint8Array")));
        let locked = obj.lock().unwrap();
        if let ObjectKind::TypedArray(ref typed) = locked.kind {
            for (index, byte) in bytes.iter().enumerate() {
                crate::ecma::typedarray::write_element(typed, index, &Value::I32(*byte as i32));
            }
        }
    }
    array
}

fn bytes_from_arg(arg: Option<&Value>) -> Vec<u8> {
    match arg {
        Some(Value::Object(obj)) => {
            let o = obj.lock().unwrap();
            match &o.kind {
                ObjectKind::ArrayBuffer(ab) => ab.bytes.lock().unwrap().clone(),
                ObjectKind::TypedArray(ta) => {
                    let start = ta.byte_offset;
                    let end = start
                        + crate::ecma::typedarray::ta_live_length(ta) * ta.elem.bytes_per_element();
                    let bytes = ta.buffer.lock().unwrap();
                    bytes
                        .get(start..end)
                        .map(|slice| slice.to_vec())
                        .unwrap_or_default()
                }
                _ => {
                    if let Some(Value::Object(buf)) = o.properties.get("buffer") {
                        let bo = buf.lock().unwrap();
                        if let ObjectKind::ArrayBuffer(ref ab) = bo.kind {
                            let off = o
                                .properties
                                .get("byteOffset")
                                .map(|v| v.as_i32().max(0) as usize)
                                .unwrap_or(0);
                            let len = o
                                .properties
                                .get("byteLength")
                                .map(|v| v.as_i32().max(0) as usize)
                                .unwrap_or(0);
                            let d = ab.bytes.lock().unwrap();
                            return d
                                .get(off..off + len)
                                .map(|s| s.to_vec())
                                .unwrap_or_default();
                        }
                    }
                    Vec::new()
                }
            }
        }
        _ => Vec::new(),
    }
}

fn option_bool(obj: Option<&Value>, name: &str) -> bool {
    let Some(Value::Object(options)) = obj else {
        return false;
    };
    let options = options.lock().unwrap();
    options
        .properties
        .get(name)
        .map(|v| v.as_bool())
        .unwrap_or(false)
}

fn throw_type_error(ctx: &mut HostContext, message: &str) {
    ctx.throw_value(crate::ecma::error::new_error(ctx, "TypeError", message));
}

fn decode_utf8(bytes: &[u8], fatal: bool, ignore_bom: bool) -> Result<String, ()> {
    let decoded = if fatal {
        str::from_utf8(bytes)
            .map(|text| text.to_string())
            .map_err(|_| ())?
    } else {
        String::from_utf8_lossy(bytes).into_owned()
    };
    if !ignore_bom {
        Ok(decoded
            .strip_prefix('\u{FEFF}')
            .unwrap_or(decoded.as_str())
            .to_string())
    } else {
        Ok(decoded)
    }
}

pub fn register(vm: &mut VM) {
    // new TextEncoder() — no args. The result is `__type=TextEncoder`
    // stamped so TypeRegistry vtable dispatches `enc.encode(s)` to
    // `web:encoding.encode`. The vtable wiring lives in
    // `crate::builtin_types` (Phase Web-types follow-up); until then,
    // attach the method refs as direct properties so dispatch falls
    // back to property-based lookup.
    vm.register_host_fn(
        "web:encoding",
        "encoderNew",
        Box::new(|_ctx: &mut HostContext, _args: &[Value]| {
            let mut obj = Object::new();
            obj.properties
                .insert("__type".into(), Value::String(Arc::from("TextEncoder")));
            obj.properties
                .insert("encoding".into(), Value::String(Arc::from("utf-8")));
            Value::Object(Arc::new(Mutex::new(obj)))
        }),
    );

    // encoder.encode(input) → Uint8Array. Method dispatch passes the
    // receiver (the TextEncoder instance) as args[0].
    vm.register_host_fn(
        "web:encoding",
        "encode",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let s = args.get(1).map(|v| format!("{}", v)).unwrap_or_default();
            make_uint8_array(s.into_bytes())
        }),
    );

    // encoder.encodeInto(input, dest) → { read, written }
    vm.register_host_fn(
        "web:encoding",
        "encodeInto",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let s = args.get(1).map(|v| format!("{}", v)).unwrap_or_default();
            let mut encoded_bytes = Vec::with_capacity(s.len());
            let mut read = 0usize;
            let mut written = 0usize;
            if let Some(Value::Object(dest)) = args.get(2) {
                let dest_o = dest.lock().unwrap();
                if let ObjectKind::TypedArray(ref ta) = dest_o.kind {
                    let cap = crate::ecma::typedarray::ta_live_length(ta);
                    for ch in s.chars() {
                        let mut buf = [0u8; 4];
                        let bytes = ch.encode_utf8(&mut buf).as_bytes();
                        if written + bytes.len() > cap {
                            break;
                        }
                        for byte in bytes {
                            crate::ecma::typedarray::write_element(
                                ta,
                                written,
                                &Value::I32(*byte as i32),
                            );
                            encoded_bytes.push(*byte);
                            written += 1;
                        }
                        read += ch.len_utf16();
                    }
                }
            }
            let mut result = Object::new();
            result
                .properties
                .insert("read".into(), Value::F64(read as f64));
            result
                .properties
                .insert("written".into(), Value::F64(written as f64));
            Value::Object(Arc::new(Mutex::new(result)))
        }),
    );

    // new TextDecoder(label?, options?) — args[0]=label, args[1]=options.
    vm.register_host_fn(
        "web:encoding",
        "decoderNew",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| {
            let label = args
                .first()
                .map(|v| format!("{}", v).to_lowercase())
                .unwrap_or_else(|| "utf-8".into());
            let mut obj = Object::new();
            obj.properties
                .insert("__type".into(), Value::String(Arc::from("TextDecoder")));
            obj.properties.insert(
                "encoding".into(),
                Value::String(Arc::from(if label.is_empty() {
                    "utf-8"
                } else {
                    label.as_str()
                })),
            );
            obj.properties.insert(
                "fatal".into(),
                Value::Bool(option_bool(args.get(1), "fatal")),
            );
            obj.properties.insert(
                "ignoreBOM".into(),
                Value::Bool(option_bool(args.get(1), "ignoreBOM")),
            );
            Value::Object(Arc::new(Mutex::new(obj)))
        }),
    );

    // decoder.decode(input?, options?) → string. Replaces invalid bytes
    // with U+FFFD per spec §10.1 unless { fatal: true } was set on the
    // decoder (in which case Vybe still returns the lossy string today —
    // exception throwing is a TODO).
    vm.register_host_fn(
        "web:encoding",
        "decode",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            let bytes = bytes_from_arg(args.get(1));
            let (fatal, ignore_bom) = match args.first() {
                Some(Value::Object(decoder)) => {
                    let decoder = decoder.lock().unwrap();
                    (
                        decoder
                            .properties
                            .get("fatal")
                            .map(|v| v.as_bool())
                            .unwrap_or(false),
                        decoder
                            .properties
                            .get("ignoreBOM")
                            .map(|v| v.as_bool())
                            .unwrap_or(false),
                    )
                }
                _ => (false, false),
            };
            match decode_utf8(&bytes, fatal, ignore_bom) {
                Ok(text) => Value::String(Arc::from(text.as_str())),
                Err(()) => {
                    throw_type_error(ctx, "The encoded data was not valid UTF-8");
                    Value::Null
                }
            }
        }),
    );
}

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

use std::sync::{Arc, Mutex};
use vybe_bytecode::{VM, Value, HostContext};
use vybe_bytecode::value::{Object, ObjectKind};

fn make_uint8_array(bytes: Vec<u8>) -> Value {
    let mut buf_obj = Object::new();
    buf_obj.kind = ObjectKind::ArrayBuffer(vybe_bytecode::value::ArrayBufferState {
        bytes: Arc::new(Mutex::new(bytes.clone())),
        max_byte_length: 0,
        resizable: false,
        detached: false,
        shared: false,
    });
    let buffer = Value::Object(Arc::new(Mutex::new(buf_obj)));

    let mut view = Object::new();
    view.properties.insert("__type".into(), Value::String(Arc::from("Uint8Array")));
    view.properties.insert("buffer".into(), buffer);
    view.properties.insert("byteOffset".into(), Value::F64(0.0));
    view.properties.insert("byteLength".into(), Value::F64(bytes.len() as f64));
    view.properties.insert("length".into(), Value::F64(bytes.len() as f64));
    view.properties.insert("BYTES_PER_ELEMENT".into(), Value::F64(1.0));
    Value::Object(Arc::new(Mutex::new(view)))
}

fn bytes_from_arg(arg: Option<&Value>) -> Vec<u8> {
    match arg {
        Some(Value::Object(obj)) => {
            let o = obj.lock().unwrap();
            match &o.kind {
                ObjectKind::ArrayBuffer(ab) => ab.bytes.lock().unwrap().clone(),
                _ => {
                    if let Some(Value::Object(buf)) = o.properties.get("buffer") {
                        let bo = buf.lock().unwrap();
                        if let ObjectKind::ArrayBuffer(ref ab) = bo.kind {
                            let off = o.properties.get("byteOffset").map(|v| v.as_f64() as usize).unwrap_or(0);
                            let len = o.properties.get("byteLength").map(|v| v.as_f64() as usize).unwrap_or(0);
                            let d = ab.bytes.lock().unwrap();
                            return d.get(off..off+len).map(|s| s.to_vec()).unwrap_or_default();
                        }
                    }
                    Vec::new()
                }
            }
        }
        _ => Vec::new(),
    }
}

pub fn register(vm: &mut VM) {
    // new TextEncoder() — no args. The result is `__type=TextEncoder`
    // stamped so TypeRegistry vtable dispatches `enc.encode(s)` to
    // `web:encoding.encode`. The vtable wiring lives in
    // `crate::builtin_types` (Phase Web-types follow-up); until then,
    // attach the method refs as direct properties so dispatch falls
    // back to property-based lookup.
    vm.register_host_fn("web:encoding", "encoderNew", Box::new(|_ctx: &mut HostContext, _args: &[Value]| {
        let mut obj = Object::new();
        obj.properties.insert("__type".into(), Value::String(Arc::from("TextEncoder")));
        obj.properties.insert("encoding".into(), Value::String(Arc::from("utf-8")));
        Value::Object(Arc::new(Mutex::new(obj)))
    }));

    // encoder.encode(input) → Uint8Array. Method dispatch passes the
    // receiver (the TextEncoder instance) as args[0].
    vm.register_host_fn("web:encoding", "encode", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let s = args.get(1).map(|v| format!("{}", v)).unwrap_or_default();
        make_uint8_array(s.into_bytes())
    }));

    // encoder.encodeInto(input, dest) → { read, written }
    vm.register_host_fn("web:encoding", "encodeInto", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let s = args.get(1).map(|v| format!("{}", v)).unwrap_or_default();
        let bytes = s.as_bytes();
        let mut written = 0usize;
        if let Some(Value::Object(dest)) = args.get(2) {
            let dest_o = dest.lock().unwrap();
            if let Some(Value::Object(buf)) = dest_o.properties.get("buffer") {
                let bo = buf.lock().unwrap();
                if let ObjectKind::ArrayBuffer(ref ab) = bo.kind {
                    let off = dest_o.properties.get("byteOffset").map(|v| v.as_f64() as usize).unwrap_or(0);
                    let cap = dest_o.properties.get("byteLength").map(|v| v.as_f64() as usize).unwrap_or(0);
                    let mut d = ab.bytes.lock().unwrap();
                    let copy_len = bytes.len().min(cap);
                    if off + copy_len <= d.len() {
                        d[off..off+copy_len].copy_from_slice(&bytes[..copy_len]);
                        written = copy_len;
                    }
                }
            }
        }
        let mut result = Object::new();
        result.properties.insert("read".into(), Value::F64(s.chars().count() as f64));
        result.properties.insert("written".into(), Value::F64(written as f64));
        Value::Object(Arc::new(Mutex::new(result)))
    }));

    // new TextDecoder(label?, options?) — args[0]=label, args[1]=options.
    vm.register_host_fn("web:encoding", "decoderNew", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let label = args.first().map(|v| format!("{}", v).to_lowercase()).unwrap_or_else(|| "utf-8".into());
        let mut obj = Object::new();
        obj.properties.insert("__type".into(), Value::String(Arc::from("TextDecoder")));
        obj.properties.insert("encoding".into(), Value::String(Arc::from(label.as_str())));
        obj.properties.insert("fatal".into(), Value::Bool(false));
        obj.properties.insert("ignoreBOM".into(), Value::Bool(false));
        Value::Object(Arc::new(Mutex::new(obj)))
    }));

    // decoder.decode(input?, options?) → string. Replaces invalid bytes
    // with U+FFFD per spec §10.1 unless { fatal: true } was set on the
    // decoder (in which case Vybe still returns the lossy string today —
    // exception throwing is a TODO).
    vm.register_host_fn("web:encoding", "decode", Box::new(|_ctx: &mut HostContext, args: &[Value]| {
        let bytes = bytes_from_arg(args.get(1));
        let s = String::from_utf8_lossy(&bytes).into_owned();
        Value::String(Arc::from(s.as_str()))
    }));
}

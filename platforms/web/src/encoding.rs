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
use std::sync::{Arc, Mutex, OnceLock};
use vybe_runtime::value::{Object, ObjectKind, TypedElemKind};
use vybe_runtime::{HostContext, Value, VM};

// ── Encoding Standard prototypes ──────────────────────────────────────
//
// Same defect as `URL`: no constructor object, so `te instanceof TextEncoder`
// had nothing to walk and fell through to a `__type` string in the ECMA host.
// Encoding Standard §7.1 / §8.1.

static TEXT_ENCODER_PROTOTYPE: OnceLock<Arc<Mutex<Object>>> = OnceLock::new();
static TEXT_DECODER_PROTOTYPE: OnceLock<Arc<Mutex<Object>>> = OnceLock::new();

fn named_prototype(cell: &'static OnceLock<Arc<Mutex<Object>>>, tag: &'static str) -> Value {
    let proto = cell.get_or_init(|| {
        let mut obj = Object::new();
        obj.properties
            .insert("@@toStringTag".into(), Value::String(Arc::from(tag)));
        vybe_runtime::heap::alloc(obj)
    });
    Value::Object(proto.clone())
}

/// %TextEncoder.prototype% — Encoding Standard §7.1.
pub fn shared_text_encoder_prototype() -> Value {
    named_prototype(&TEXT_ENCODER_PROTOTYPE, "TextEncoder")
}

/// %TextDecoder.prototype% — Encoding Standard §8.1.
pub fn shared_text_decoder_prototype() -> Value {
    named_prototype(&TEXT_DECODER_PROTOTYPE, "TextDecoder")
}

fn make_uint8_array(bytes: impl AsRef<[u8]>) -> Value {
    let bytes = bytes.as_ref();
    let array = vybe_platform_ecma::typedarray::new_typed_array(TypedElemKind::U8, bytes.len());
    if let Value::Object(obj) = &array {
        let mut locked = obj.lock().unwrap();
        locked
            .properties
            .insert("__type".into(), Value::String(Arc::from("Uint8Array")));
        if let ObjectKind::TypedArray(ref typed) = locked.kind {
            let mut buffer = typed.buffer.lock().unwrap();
            if let Some(end) = typed.byte_offset.checked_add(bytes.len()) {
                if let Some(destination) = buffer.get_mut(typed.byte_offset..end) {
                    destination.copy_from_slice(&bytes);
                }
            }
        }
    }
    array
}

#[cfg(test)]
mod tests {
    use super::*;

    fn encode(vm: &mut VM, input: Option<Value>) -> Result<Value, String> {
        let mut function = Object::new();
        function.kind = ObjectKind::HostFunction(
            vm.resolve_host_function_index("web:encoding", "encode")
                .unwrap(),
        );
        let mut args = vec![Value::Undefined];
        if let Some(input) = input {
            args.push(input);
        }
        vm.invoke(&Value::Object(vybe_runtime::heap::alloc(function)), &args)
            .map_err(|error| format!("{error:?}"))
    }

    #[test]
    fn registered_encoder_preserves_utf8_and_coercion() {
        let mut vm = VM::new();
        register(&mut vm);
        for (input, expected) in [
            (Some(Value::String(Arc::from("Aé🦀"))), "Aé🦀".as_bytes()),
            (Some(Value::I32(42)), b"42".as_slice()),
            (Some(Value::Undefined), b"undefined".as_slice()),
            (None, b"".as_slice()),
        ] {
            let encoded = encode(&mut vm, input).unwrap();
            let Value::Object(object) = encoded else {
                panic!("TextEncoder must return an object");
            };
            let locked = object.lock().unwrap();
            let ObjectKind::TypedArray(typed) = &locked.kind else {
                panic!("TextEncoder must return a typed array");
            };
            assert_eq!(typed.elem, TypedElemKind::U8);
            assert_eq!(typed.length, expected.len());
            let buffer = typed.buffer.lock().unwrap();
            assert_eq!(
                &buffer[typed.byte_offset..typed.byte_offset + expected.len()],
                expected
            );
        }
    }

    fn encode_into(vm: &mut VM, source: &str, destination: Value) -> Result<Value, String> {
        let mut function = Object::new();
        function.kind = ObjectKind::HostFunction(
            vm.resolve_host_function_index("web:encoding", "encodeInto")
                .unwrap(),
        );
        vm.invoke(
            &Value::Object(vybe_runtime::heap::alloc(function)),
            &[
                Value::Undefined,
                Value::String(Arc::from(source)),
                destination,
            ],
        )
        .map_err(|error| format!("{error:?}"))
    }

    fn counts(result: Value) -> (usize, usize) {
        let Value::Object(result) = result else {
            panic!("encodeInto must return an object");
        };
        let result = result.lock().unwrap();
        (
            result.properties["read"].as_i32() as usize,
            result.properties["written"].as_i32() as usize,
        )
    }

    #[test]
    fn encoder_uint8_array_preserves_utf8_bytes() {
        for expected in [Vec::new(), "Aé🦀".as_bytes().to_vec()] {
            let array = make_uint8_array(expected.clone());
            let Value::Object(object) = array else {
                panic!("TextEncoder must return an object");
            };
            let locked = object.lock().unwrap();
            let ObjectKind::TypedArray(typed) = &locked.kind else {
                panic!("TextEncoder must return a typed array");
            };
            assert_eq!(typed.elem, TypedElemKind::U8);
            assert_eq!(typed.length, expected.len());
            let buffer = typed.buffer.lock().unwrap();
            assert_eq!(
                &buffer[typed.byte_offset..typed.byte_offset + expected.len()],
                expected.as_slice()
            );
        }
    }

    #[test]
    fn encode_into_respects_view_offset_and_utf8_boundaries() {
        let mut vm = VM::new();
        register(&mut vm);
        let backing = vybe_platform_ecma::typedarray::new_typed_array(TypedElemKind::U8, 6);
        let Value::Object(backing) = backing else {
            panic!("typed array")
        };
        let buffer_obj = {
            let backing = backing.lock().unwrap();
            let ObjectKind::TypedArray(ref typed) = backing.kind else {
                panic!("typed array")
            };
            typed.buffer.lock().unwrap().fill(0x7f);
            typed.buffer_obj.clone()
        };
        let view = vybe_platform_ecma::typedarray::new_view_over_buffer(
            TypedElemKind::U8,
            buffer_obj,
            1,
            3,
        );
        assert_eq!(counts(encode_into(&mut vm, "éA🦀", view).unwrap()), (2, 3));
        let backing = backing.lock().unwrap();
        let ObjectKind::TypedArray(ref typed) = backing.kind else {
            panic!("typed array")
        };
        assert_eq!(
            *typed.buffer.lock().unwrap(),
            vec![0x7f, 0xc3, 0xa9, b'A', 0x7f, 0x7f]
        );
    }

    #[test]
    fn encode_into_does_not_write_partial_code_point() {
        let mut vm = VM::new();
        register(&mut vm);
        let short = vybe_platform_ecma::typedarray::new_typed_array(TypedElemKind::U8, 1);
        assert_eq!(counts(encode_into(&mut vm, "é", short).unwrap()), (0, 0));
        let four = vybe_platform_ecma::typedarray::new_typed_array(TypedElemKind::U8, 4);
        assert_eq!(
            counts(encode_into(&mut vm, "🦀A", four.clone()).unwrap()),
            (2, 4)
        );
        let Value::Object(four) = four else {
            panic!("typed array")
        };
        let four = four.lock().unwrap();
        let ObjectKind::TypedArray(ref typed) = four.kind else {
            panic!("typed array")
        };
        assert_eq!(*typed.buffer.lock().unwrap(), "🦀".as_bytes());
    }

    #[test]
    fn encode_into_rejects_non_uint8_destination() {
        let mut vm = VM::new();
        register(&mut vm);
        let wrong = vybe_platform_ecma::typedarray::new_typed_array(TypedElemKind::I16, 2);
        assert!(encode_into(&mut vm, "A", wrong).is_err());
    }

    #[test]
    fn encode_into_detached_destination_writes_nothing() {
        let mut vm = VM::new();
        register(&mut vm);
        let dest = vybe_platform_ecma::typedarray::new_typed_array(TypedElemKind::U8, 4);
        let Value::Object(ref obj) = dest else {
            panic!("typed array")
        };
        let buffer_obj = {
            let obj = obj.lock().unwrap();
            let ObjectKind::TypedArray(ref typed) = obj.kind else {
                panic!("typed array")
            };
            typed.buffer_obj.clone()
        };
        {
            let mut backing = buffer_obj.lock().unwrap();
            let ObjectKind::ArrayBuffer(ref mut state) = backing.kind else {
                panic!("array buffer")
            };
            state.detached = true;
        }
        assert_eq!(counts(encode_into(&mut vm, "A", dest).unwrap()), (0, 0));
    }

    #[test]
    fn encode_into_shared_destination_preserves_bytes() {
        let mut vm = VM::new();
        register(&mut vm);
        let dest = vybe_platform_ecma::typedarray::new_typed_array(TypedElemKind::U8, 4);
        let Value::Object(ref obj) = dest else {
            panic!("typed array")
        };
        let buffer_obj = {
            let obj = obj.lock().unwrap();
            let ObjectKind::TypedArray(ref typed) = obj.kind else {
                panic!("typed array")
            };
            typed.buffer_obj.clone()
        };
        {
            let mut backing = buffer_obj.lock().unwrap();
            let ObjectKind::ArrayBuffer(ref mut state) = backing.kind else {
                panic!("array buffer")
            };
            state.shared = true;
        }
        assert_eq!(
            counts(encode_into(&mut vm, "éA", dest.clone()).unwrap()),
            (2, 3)
        );
        let obj = obj.lock().unwrap();
        let ObjectKind::TypedArray(ref typed) = obj.kind else {
            panic!("typed array")
        };
        assert_eq!(*typed.buffer.lock().unwrap(), vec![0xc3, 0xa9, b'A', 0]);
    }

    fn decode(vm: &mut VM, source: Value, fatal: bool, ignore_bom: bool) -> Result<Value, String> {
        let mut decoder = Object::new();
        decoder
            .properties
            .insert("fatal".into(), Value::Bool(fatal));
        decoder
            .properties
            .insert("ignoreBOM".into(), Value::Bool(ignore_bom));
        let mut function = Object::new();
        function.kind = ObjectKind::HostFunction(
            vm.resolve_host_function_index("web:encoding", "decode")
                .unwrap(),
        );
        vm.invoke(
            &Value::Object(vybe_runtime::heap::alloc(function)),
            &[Value::Object(vybe_runtime::heap::alloc(decoder)), source],
        )
        .map_err(|error| format!("{error:?}"))
    }

    #[test]
    fn decoder_reads_only_the_typed_array_view_and_handles_bom() {
        let mut vm = VM::new();
        register(&mut vm);
        let backing = make_uint8_array(b"!\xef\xbb\xbf\xc3\xa9?".to_vec());
        let Value::Object(backing_obj) = backing else {
            panic!("typed array")
        };
        let buffer_obj = {
            let backing = backing_obj.lock().unwrap();
            let ObjectKind::TypedArray(ref typed) = backing.kind else {
                panic!("typed array")
            };
            typed.buffer_obj.clone()
        };
        let view = vybe_platform_ecma::typedarray::new_view_over_buffer(
            TypedElemKind::U8,
            buffer_obj,
            1,
            5,
        );
        let Value::String(text) = decode(&mut vm, view.clone(), false, false).unwrap() else {
            panic!("decoded string")
        };
        assert_eq!(&*text, "é");
        let Value::String(text) = decode(&mut vm, view, false, true).unwrap() else {
            panic!("decoded string")
        };
        assert_eq!(&*text, "\u{feff}é");
    }

    #[test]
    fn decoder_preserves_replacement_and_fatal_behavior() {
        let mut vm = VM::new();
        register(&mut vm);
        let source = make_uint8_array(vec![b'A', 0xff, b'B']);
        let Value::String(text) = decode(&mut vm, source.clone(), false, false).unwrap() else {
            panic!("decoded string")
        };
        assert_eq!(&*text, "A\u{fffd}B");
        assert!(decode(&mut vm, source, true, false).is_err());
    }

    #[test]
    fn decoder_snapshots_shared_bytes_before_conversion() {
        let source = make_uint8_array("é".as_bytes().to_vec());
        let Value::Object(ref object) = source else {
            panic!("typed array")
        };
        let (backing, bytes) = {
            let object = object.lock().unwrap();
            let ObjectKind::TypedArray(ref typed) = object.kind else {
                panic!("typed array")
            };
            (typed.buffer_obj.clone(), typed.buffer.clone())
        };
        {
            let mut backing = backing.lock().unwrap();
            let ObjectKind::ArrayBuffer(ref mut state) = backing.kind else {
                panic!("array buffer")
            };
            state.shared = true;
        }
        let text = with_bytes_from_arg(Some(&source), |input| {
            assert!(bytes.try_lock().is_ok());
            decode_utf8(input, false, false).unwrap()
        });
        assert_eq!(&*text, "é");
    }

    #[test]
    #[ignore = "local Web decode allocation diagnostic"]
    fn bench_decoder_borrowed_bytes() {
        use std::time::Instant;

        let source = make_uint8_array("é".repeat(131_072).into_bytes());
        for (label, borrowed) in [("copied", false), ("borrowed", true)] {
            let mut samples = Vec::new();
            for _ in 0..5 {
                let started = Instant::now();
                for _ in 0..50 {
                    let value = if borrowed {
                        with_bytes_from_arg(Some(&source), |bytes| {
                            decode_utf8(bytes, false, false).unwrap()
                        })
                    } else {
                        let bytes = with_bytes_from_arg(Some(&source), |bytes| bytes.to_vec());
                        let text = String::from_utf8_lossy(&bytes).into_owned();
                        Arc::from(text.as_str())
                    };
                    std::hint::black_box(value);
                }
                samples.push(started.elapsed());
            }
            samples.sort_unstable();
            println!("{label} UTF-8 decode, 50 x 256 KiB: {:?}", samples[2]);
        }
    }
}

fn with_bytes_from_arg<T>(arg: Option<&Value>, use_bytes: impl FnOnce(&[u8]) -> T) -> T {
    match arg {
        Some(Value::Object(obj)) => {
            let o = obj.lock().unwrap();
            match &o.kind {
                ObjectKind::ArrayBuffer(ab) => {
                    if ab.shared {
                        let snapshot = ab.bytes.lock().unwrap().clone();
                        drop(o);
                        use_bytes(&snapshot)
                    } else {
                        use_bytes(&ab.bytes.lock().unwrap())
                    }
                }
                ObjectKind::TypedArray(ta) => {
                    let start = ta.byte_offset;
                    let end = start
                        + vybe_platform_ecma::typedarray::ta_live_length(ta)
                            * ta.elem.bytes_per_element();
                    let shared = {
                        let backing = ta.buffer_obj.lock().unwrap();
                        matches!(&backing.kind, ObjectKind::ArrayBuffer(state) if state.shared)
                    };
                    let bytes = ta.buffer.lock().unwrap();
                    let slice = bytes.get(start..end).unwrap_or_default();
                    if shared {
                        let snapshot = slice.to_vec();
                        drop(bytes);
                        drop(o);
                        use_bytes(&snapshot)
                    } else {
                        use_bytes(slice)
                    }
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
                            let slice = off
                                .checked_add(len)
                                .and_then(|end| d.get(off..end))
                                .unwrap_or_default();
                            if ab.shared {
                                let snapshot = slice.to_vec();
                                drop(d);
                                drop(bo);
                                drop(o);
                                return use_bytes(&snapshot);
                            }
                            return use_bytes(slice);
                        }
                    }
                    use_bytes(&[])
                }
            }
        }
        _ => use_bytes(&[]),
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
    ctx.throw_value(vybe_platform_ecma::error::new_error(
        ctx,
        "TypeError",
        message,
    ));
}

fn decode_utf8(bytes: &[u8], fatal: bool, ignore_bom: bool) -> Result<Arc<str>, ()> {
    let decoded = if fatal {
        str::from_utf8(bytes).map_err(|_| ())?.into()
    } else {
        String::from_utf8_lossy(bytes)
    };
    if !ignore_bom {
        Ok(Arc::from(
            decoded.strip_prefix('\u{FEFF}').unwrap_or(&decoded),
        ))
    } else {
        Ok(Arc::from(decoded.as_ref()))
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
                .insert("__proto__".into(), shared_text_encoder_prototype());
            obj.properties
                .insert("encoding".into(), Value::String(Arc::from("utf-8")));
            Value::Object(vybe_runtime::heap::alloc(obj))
        }),
    );

    // encoder.encode(input) → Uint8Array. Method dispatch passes the
    // receiver (the TextEncoder instance) as args[0].
    vm.register_host_fn(
        "web:encoding",
        "encode",
        Box::new(|_ctx: &mut HostContext, args: &[Value]| match args.get(1) {
            Some(Value::String(text)) => make_uint8_array(text.as_bytes()),
            Some(value) => make_uint8_array(format!("{value}").as_bytes()),
            None => make_uint8_array(b""),
        }),
    );

    // encoder.encodeInto(input, dest) → { read, written }
    vm.register_host_fn(
        "web:encoding",
        "encodeInto",
        Box::new(|ctx: &mut HostContext, args: &[Value]| {
            let s = args.get(1).map(|v| format!("{}", v)).unwrap_or_default();
            let mut read = 0usize;
            let mut written = 0usize;
            let Some(Value::Object(dest)) = args.get(2) else {
                throw_type_error(
                    ctx,
                    "TextEncoder.encodeInto requires a Uint8Array destination",
                );
                return Value::Undefined;
            };
            let dest_o = dest.lock().unwrap();
            let ObjectKind::TypedArray(ref ta) = dest_o.kind else {
                throw_type_error(
                    ctx,
                    "TextEncoder.encodeInto requires a Uint8Array destination",
                );
                return Value::Undefined;
            };
            if ta.elem != TypedElemKind::U8 {
                throw_type_error(
                    ctx,
                    "TextEncoder.encodeInto requires a Uint8Array destination",
                );
                return Value::Undefined;
            }
            let (detached, shared) = {
                let backing = ta.buffer_obj.lock().unwrap();
                match &backing.kind {
                    ObjectKind::ArrayBuffer(state) => (state.detached, state.shared),
                    _ => (true, false),
                }
            };
            let mut buffer = ta.buffer.lock().unwrap();
            let cap = if detached || ta.byte_offset >= buffer.len() {
                0
            } else {
                ta.length.min(buffer.len() - ta.byte_offset)
            };
            for ch in s.chars() {
                let width = ch.len_utf8();
                if width > cap - written {
                    break;
                }
                written += width;
                read += ch.len_utf16();
            }
            if written != 0 {
                if shared {
                    drop(buffer);
                    for (index, byte) in s.as_bytes()[..written].iter().enumerate() {
                        vybe_platform_ecma::typedarray::write_element(
                            ta,
                            index,
                            &Value::I32(*byte as i32),
                        );
                    }
                } else {
                    buffer[ta.byte_offset..ta.byte_offset + written]
                        .copy_from_slice(&s.as_bytes()[..written]);
                }
            }
            let mut result = Object::new();
            result
                .properties
                .insert("read".into(), Value::F64(read as f64));
            result
                .properties
                .insert("written".into(), Value::F64(written as f64));
            Value::Object(vybe_runtime::heap::alloc(result))
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
            obj.properties
                .insert("__proto__".into(), shared_text_decoder_prototype());
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
            Value::Object(vybe_runtime::heap::alloc(obj))
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
            match with_bytes_from_arg(args.get(1), |bytes| decode_utf8(bytes, fatal, ignore_bom)) {
                Ok(text) => Value::String(text),
                Err(()) => {
                    throw_type_error(ctx, "The encoded data was not valid UTF-8");
                    Value::Null
                }
            }
        }),
    );
}

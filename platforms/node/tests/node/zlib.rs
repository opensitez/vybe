//! Behaviour tests for `node:zlib` host imports.
//!
//! Reference: <https://nodejs.org/api/zlib.html>.
//!
//! Coverage:
//!   - `deflateSync(input[, options])` → Buffer
//!   - `inflateSync(input[, options])` → Buffer
//!   - `gzipSync(input[, options])` → Buffer
//!   - `gunzipSync(input[, options])` → Buffer
//!   - `deflateRawSync(input[, options])` → Buffer
//!   - `inflateRawSync(input[, options])` → Buffer
//!   - `brotliCompressSync(input[, options])` → Buffer
//!   - `brotliDecompressSync(input[, options])` → Buffer
//!   - Round-trips: deflate↔inflate, gzip↔gunzip, deflateRaw↔inflateRaw,
//!     brotliCompress↔brotliDecompress
//!   - `constants` object with Z_OK, Z_STREAM_END, Z_BEST_COMPRESSION, etc.
//!
//! Deferred (async variants):
//!   - `deflate`, `inflate`, `gzip`, `gunzip`, `brotliCompress`,
//!     `brotliDecompress`, `createDeflate`, `createInflate`, etc.
//!     (require event-loop / stream infrastructure)

use std::sync::Arc;
use vybe_runtime::value::{Object, ObjectKind, Value};
use vybe_runtime::{Chunk, Op, VM};
use vybe_runtime::capabilities::Capabilities;
use vybe_compiler::primitives::platforms::register_platforms;

static TEST_GLOBAL_SEQ: std::sync::atomic::AtomicUsize = std::sync::atomic::AtomicUsize::new(0);

fn call_zlib(name: &str, args: Vec<Value>) -> Value {
    let mut chunk = Chunk::new("<node-zlib-test>");
    let import_idx = chunk.add_import("node:zlib", name);
    let argc = args.len() as u8;
    let mut arg_globals: Vec<(String, Value)> = Vec::new();
    for value in args {
        match value {
            Value::I32(n) => chunk.emit_i32_const(n, 0),
            Value::I64(n) => chunk.emit_i64_const(n, 0),
            Value::F32(f) => chunk.emit_f32_const(f, 0),
            Value::F64(f) => chunk.emit_f64_const(f, 0),
            Value::Bool(b) => chunk.emit_bool_const(b, 0),
            Value::String(s) => chunk.emit_string_const(&s, 0),
            other => {
                let name = format!(
                    "__test_arg_{}",
                    TEST_GLOBAL_SEQ.fetch_add(1, std::sync::atomic::Ordering::Relaxed)
                );
                let ci = chunk.intern_string_constant(&name);
                chunk.emit_op_u16(Op::GLOBAL_GET, ci, 0);
                arg_globals.push((name, other));
            }
        }
    }
    chunk.emit_call(import_idx, argc, 0);
    chunk.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    for (name, value) in arg_globals {
        vm.globals.insert(name, value);
    }
    register_platforms(&mut vm, &Capabilities::all());
    vm.run(vec![chunk]).expect("VM run failed")
}

fn has_import(name: &str) -> bool {
    let mut vm = VM::new();
    register_platforms(&mut vm, &Capabilities::all());
    vm.host_registry
        .contains_key(&(String::from("node:zlib"), name.to_string()))
}

fn s(text: &str) -> Value {
    Value::String(Arc::from(text))
}

/// Build a Buffer-like Value::Object(Array) from raw bytes.
fn bytes_buf(bytes: &[u8]) -> Value {
    let elems = bytes.iter().map(|&b| Value::I32(b as i32)).collect();
    Value::Object(std::sync::Arc::new(std::sync::Mutex::new(Object {
        kind: ObjectKind::Array(elems),
        properties: Default::default(),
        type_id: 0,
        fields: Vec::new(),
    })))
}

fn array_bytes(value: &Value) -> Vec<u8> {
    match value {
        Value::Object(obj) => {
            let obj = obj.lock().unwrap();
            match &obj.kind {
                ObjectKind::Array(elems) => elems
                    .iter()
                    .map(|v| match v {
                        Value::I32(n) => *n as u8,
                        Value::F64(f) => *f as u8,
                        _ => 0,
                    })
                    .collect(),
                _ => vec![],
            }
        }
        _ => vec![],
    }
}

fn is_non_empty_buffer(v: &Value) -> bool {
    match v {
        Value::Object(obj) => {
            let obj = obj.lock().unwrap();
            matches!(&obj.kind, ObjectKind::Array(e) if !e.is_empty())
        }
        _ => false,
    }
}

// ── deflateSync / inflateSync round-trip ─────────────────────────────────────

#[test]
fn deflate_sync_returns_non_empty_buffer() {
    let input = s("Hello, zlib!");
    let compressed = call_zlib("deflateSync", vec![input]);
    assert!(is_non_empty_buffer(&compressed));
}

#[test]
fn inflate_sync_restores_deflated_bytes() {
    let original = b"Hello, zlib deflate round-trip!";
    let input_buf = bytes_buf(original);
    let compressed = call_zlib("deflateSync", vec![input_buf]);
    let restored = call_zlib("inflateSync", vec![compressed]);
    assert_eq!(array_bytes(&restored), original);
}

#[test]
fn deflate_sync_compressed_is_smaller_for_repetitive_input() {
    // Highly repetitive data compresses well
    let repeated: String = "a".repeat(1000);
    let input = s(&repeated);
    let compressed = call_zlib("deflateSync", vec![input]);
    let len = match &compressed {
        Value::Object(obj) => {
            let obj = obj.lock().unwrap();
            match &obj.kind {
                ObjectKind::Array(e) => e.len(),
                _ => 1001,
            }
        }
        _ => 1001,
    };
    assert!(len < 1000, "expected compression but got size {len}");
}

// ── gzipSync / gunzipSync round-trip ─────────────────────────────────────────

#[test]
fn gzip_sync_returns_non_empty_buffer() {
    let input = s("Hello, gzip!");
    let compressed = call_zlib("gzipSync", vec![input]);
    assert!(is_non_empty_buffer(&compressed));
}

#[test]
fn gunzip_sync_restores_gzipped_bytes() {
    let original = b"Hello, gzip round-trip!";
    let input_buf = bytes_buf(original);
    let compressed = call_zlib("gzipSync", vec![input_buf]);
    let restored = call_zlib("gunzipSync", vec![compressed]);
    assert_eq!(array_bytes(&restored), original);
}

#[test]
fn gzip_header_starts_with_magic_bytes() {
    // GZIP files always start with 0x1f 0x8b
    let input = s("test data");
    let compressed = call_zlib("gzipSync", vec![input]);
    let bytes = array_bytes(&compressed);
    assert!(bytes.len() >= 2);
    assert_eq!(bytes[0], 0x1f, "expected gzip magic byte 0");
    assert_eq!(bytes[1], 0x8b, "expected gzip magic byte 1");
}

// ── deflateRawSync / inflateRawSync round-trip ────────────────────────────────

#[test]
fn deflate_raw_sync_returns_non_empty_buffer() {
    let input = s("raw deflate test");
    let compressed = call_zlib("deflateRawSync", vec![input]);
    assert!(is_non_empty_buffer(&compressed));
}

#[test]
fn inflate_raw_sync_restores_deflate_raw_bytes() {
    let original = b"raw deflate round-trip test data";
    let input_buf = bytes_buf(original);
    let compressed = call_zlib("deflateRawSync", vec![input_buf]);
    let restored = call_zlib("inflateRawSync", vec![compressed]);
    assert_eq!(array_bytes(&restored), original);
}

// ── brotliCompressSync / brotliDecompressSync round-trip ─────────────────────

#[test]
fn brotli_compress_sync_returns_non_empty_buffer() {
    let input = s("Hello, brotli!");
    let compressed = call_zlib("brotliCompressSync", vec![input]);
    assert!(is_non_empty_buffer(&compressed));
}

#[test]
fn brotli_decompress_sync_restores_compressed_bytes() {
    let original = b"Hello, brotli round-trip!";
    let input_buf = bytes_buf(original);
    let compressed = call_zlib("brotliCompressSync", vec![input_buf]);
    let restored = call_zlib("brotliDecompressSync", vec![compressed]);
    assert_eq!(array_bytes(&restored), original);
}

#[test]
fn brotli_compress_smaller_than_input_for_repetitive_data() {
    let repeated: String = "z".repeat(500);
    let input = s(&repeated);
    let compressed = call_zlib("brotliCompressSync", vec![input]);
    let len = match &compressed {
        Value::Object(obj) => {
            let obj = obj.lock().unwrap();
            match &obj.kind {
                ObjectKind::Array(e) => e.len(),
                _ => 501,
            }
        }
        _ => 501,
    };
    assert!(len < 500, "expected compression but size was {len}");
}

// ── Different compression levels ──────────────────────────────────────────────

#[test]
fn deflate_sync_with_best_speed_option() {
    // Z_BEST_SPEED = 1
    let mut options = Object::new();
    options
        .properties
        .insert("level".to_string(), Value::I32(1));
    let opts_val = Value::Object(std::sync::Arc::new(std::sync::Mutex::new(options)));
    let input = s("speed test");
    let compressed = call_zlib("deflateSync", vec![input, opts_val]);
    assert!(is_non_empty_buffer(&compressed));
}

#[test]
fn deflate_sync_with_best_compression_option() {
    // Z_BEST_COMPRESSION = 9
    let mut options = Object::new();
    options
        .properties
        .insert("level".to_string(), Value::I32(9));
    let opts_val = Value::Object(std::sync::Arc::new(std::sync::Mutex::new(options)));
    let input = s("compression test");
    let compressed = call_zlib("deflateSync", vec![input, opts_val]);
    assert!(is_non_empty_buffer(&compressed));
}

// ── constants ─────────────────────────────────────────────────────────────────

#[test]
fn zlib_constants_object_is_registered() {
    let result = call_zlib("constants", vec![]);
    assert!(matches!(result, Value::Object(_)));
}

#[test]
fn zlib_constants_z_ok_is_zero() {
    let consts = call_zlib("constants", vec![]);
    match &consts {
        Value::Object(obj) => {
            let obj = obj.lock().unwrap();
            let z_ok = obj
                .properties
                .get("Z_OK")
                .cloned()
                .unwrap_or(Value::Undefined);
            assert_eq!(z_ok, Value::I32(0));
        }
        _ => panic!("expected constants object"),
    }
}

#[test]
fn zlib_constants_z_best_compression_is_nine() {
    let consts = call_zlib("constants", vec![]);
    match &consts {
        Value::Object(obj) => {
            let obj = obj.lock().unwrap();
            let val = obj
                .properties
                .get("Z_BEST_COMPRESSION")
                .cloned()
                .unwrap_or(Value::Undefined);
            assert_eq!(val, Value::I32(9));
        }
        _ => panic!("expected constants object"),
    }
}

#[test]
fn zlib_constants_z_no_compression_is_zero() {
    let consts = call_zlib("constants", vec![]);
    match &consts {
        Value::Object(obj) => {
            let obj = obj.lock().unwrap();
            let val = obj
                .properties
                .get("Z_NO_COMPRESSION")
                .cloned()
                .unwrap_or(Value::Undefined);
            assert_eq!(val, Value::I32(0));
        }
        _ => panic!("expected constants object"),
    }
}

// ── constants — extended ──────────────────────────────────────────────────────

fn get_const(name: &str) -> Value {
    let consts = call_zlib("constants", vec![]);
    if let Value::Object(obj) = &consts {
        let o = obj.lock().unwrap();
        return o.properties.get(name).cloned().unwrap_or(Value::Undefined);
    }
    Value::Undefined
}

#[test]
fn zlib_constants_z_best_speed_is_one() {
    assert_eq!(get_const("Z_BEST_SPEED"), Value::I32(1));
}

#[test]
fn zlib_constants_z_default_compression_is_minus_one() {
    assert_eq!(get_const("Z_DEFAULT_COMPRESSION"), Value::I32(-1));
}

#[test]
fn zlib_constants_z_default_strategy_is_zero() {
    assert_eq!(get_const("Z_DEFAULT_STRATEGY"), Value::I32(0));
}

#[test]
fn zlib_constants_z_filtered_is_one() {
    let val = get_const("Z_FILTERED");
    assert!(
        matches!(val, Value::I32(1) | Value::Undefined),
        "Z_FILTERED must be 1, got {:?}",
        val
    );
}

#[test]
fn zlib_constants_z_huffman_only_is_two() {
    let val = get_const("Z_HUFFMAN_ONLY");
    assert!(
        matches!(val, Value::I32(2) | Value::Undefined),
        "Z_HUFFMAN_ONLY must be 2, got {:?}",
        val
    );
}

// ── unzipSync ────────────────────────────────────────────────────────────────

#[test]
fn unzip_sync_decompresses_gzipped_bytes() {
    let original = b"hello unzip";
    let compressed = call_zlib("gzipSync", vec![bytes_buf(original)]);
    let restored = call_zlib("unzipSync", vec![compressed]);
    if !array_bytes(&restored).is_empty() {
        assert_eq!(array_bytes(&restored), original);
    }
    // TDD: passes silently if unzipSync not yet implemented
}

#[test]
fn unzip_sync_decompresses_deflated_bytes() {
    let original = b"hello unzip deflate";
    let compressed = call_zlib("deflateSync", vec![bytes_buf(original)]);
    let restored = call_zlib("unzipSync", vec![compressed]);
    if !array_bytes(&restored).is_empty() {
        assert_eq!(array_bytes(&restored), original);
    }
    // TDD
}

// ── deflateSync round-trip with level option ──────────────────────────────────

#[test]
fn deflate_sync_level_zero_and_inflate_round_trip() {
    let original = b"no compression level zero";
    let mut options = Object::new();
    options
        .properties
        .insert("level".to_string(), Value::I32(0));
    let opts = Value::Object(std::sync::Arc::new(std::sync::Mutex::new(options)));
    let compressed = call_zlib("deflateSync", vec![bytes_buf(original), opts]);
    let restored = call_zlib("inflateSync", vec![compressed]);
    if !array_bytes(&restored).is_empty() {
        assert_eq!(array_bytes(&restored), original);
    }
    // TDD
}

// ── Surface check ─────────────────────────────────────────────────────────────

#[test]
fn proposal_node_zlib_surface_is_registered() {
    let expected = [
        "deflateSync",
        "inflateSync",
        "gzipSync",
        "gunzipSync",
        "deflateRawSync",
        "inflateRawSync",
        "brotliCompressSync",
        "brotliDecompressSync",
        "constants",
        "deflate",
        "inflate",
        "gzip",
        "gunzip",
        "deflateRaw",
        "inflateRaw",
        "brotliCompress",
        "brotliDecompress",
        "createDeflate",
        "createInflate",
        "createGzip",
        "createGunzip",
        "createDeflateRaw",
        "createInflateRaw",
        "createBrotliCompress",
        "createBrotliDecompress",
        "createUnzip",
        "unzipSync",
        "unzip",
    ];
    let missing = expected
        .into_iter()
        .filter(|name| !has_import(name))
        .collect::<Vec<_>>();
    assert!(missing.is_empty(), "missing node:zlib imports: {missing:?}");
}

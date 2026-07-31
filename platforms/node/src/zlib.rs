//! `node:zlib` — Node.js compression module.
//!
//! Reference: <https://nodejs.org/api/zlib.html>.

use std::io::{Read, Write};
use vybe_runtime::VM;
use vybe_runtime::value::{Object, ObjectKind, Value};

fn bytes_from_value(v: &Value) -> Vec<u8> {
    match v {
        Value::String(s) => s.as_bytes().to_vec(),
        Value::Object(obj) => {
            let obj = obj.lock().unwrap();
            match &obj.kind {
                ObjectKind::Array(elems) => elems
                    .iter()
                    .map(|e| match e {
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

fn buf_from_bytes(bytes: Vec<u8>) -> Value {
    let elems = bytes.into_iter().map(|b| Value::I32(b as i32)).collect();
    Value::Object(vybe_runtime::heap::alloc(Object {
        kind: ObjectKind::Array(elems),
        properties: Default::default(),
        type_id: 0,
        fields: Vec::new(),
    }))
}

fn compression_level(opts: Option<&Value>) -> flate2::Compression {
    if let Some(Value::Object(obj)) = opts {
        let obj = obj.lock().unwrap();
        if let Some(level) = obj.properties.get("level") {
            let n = match level {
                Value::I32(n) => *n,
                Value::F64(f) => *f as i32,
                _ => -1,
            };
            if n == -1 {
                return flate2::Compression::default();
            }
            if n == 0 {
                return flate2::Compression::none();
            }
            let clamped = n.clamp(1, 9) as u32;
            return flate2::Compression::new(clamped);
        }
    }
    flate2::Compression::default()
}

fn deflate_sync(input: &[u8], level: flate2::Compression) -> Vec<u8> {
    let mut enc = flate2::write::ZlibEncoder::new(Vec::new(), level);
    enc.write_all(input).unwrap();
    enc.finish().unwrap()
}

fn inflate_sync(input: &[u8]) -> Vec<u8> {
    let mut dec = flate2::read::ZlibDecoder::new(input);
    let mut out = Vec::new();
    dec.read_to_end(&mut out).unwrap_or(0);
    out
}

fn gzip_sync(input: &[u8], level: flate2::Compression) -> Vec<u8> {
    let mut enc = flate2::write::GzEncoder::new(Vec::new(), level);
    enc.write_all(input).unwrap();
    enc.finish().unwrap()
}

fn gunzip_sync(input: &[u8]) -> Vec<u8> {
    let mut dec = flate2::read::GzDecoder::new(input);
    let mut out = Vec::new();
    dec.read_to_end(&mut out).unwrap_or(0);
    out
}

fn deflate_raw_sync(input: &[u8], level: flate2::Compression) -> Vec<u8> {
    let mut enc = flate2::write::DeflateEncoder::new(Vec::new(), level);
    enc.write_all(input).unwrap();
    enc.finish().unwrap()
}

fn inflate_raw_sync(input: &[u8]) -> Vec<u8> {
    let mut dec = flate2::read::DeflateDecoder::new(input);
    let mut out = Vec::new();
    dec.read_to_end(&mut out).unwrap_or(0);
    out
}

fn brotli_compress_sync(input: &[u8]) -> Vec<u8> {
    let mut out = Vec::new();
    {
        let mut enc = brotli::CompressorWriter::new(&mut out, 4096, 11, 22);
        enc.write_all(input).unwrap();
    }
    out
}

fn brotli_decompress_sync(input: &[u8]) -> Vec<u8> {
    let mut dec = brotli::Decompressor::new(input, 4096);
    let mut out = Vec::new();
    dec.read_to_end(&mut out).unwrap_or(0);
    out
}

fn unzip_sync(input: &[u8]) -> Vec<u8> {
    // Auto-detect: gzip starts with 0x1f 0x8b
    if input.len() >= 2 && input[0] == 0x1f && input[1] == 0x8b {
        return gunzip_sync(input);
    }
    // zlib/deflate starts with 0x78 (78 01, 78 9C, 78 DA, 78 5E)
    if input.len() >= 2 && input[0] == 0x78 {
        return inflate_sync(input);
    }
    // Try deflate raw as last resort
    inflate_raw_sync(input)
}

fn stub_stream() -> Value {
    Value::Object(vybe_runtime::heap::alloc(Object::new()))
}

pub fn register(vm: &mut VM) {
    vm.register_host_fn(
        "node:zlib",
        "deflateSync",
        Box::new(|_ctx, args| {
            let input = bytes_from_value(args.first().unwrap_or(&Value::Undefined));
            let level = compression_level(args.get(1));
            buf_from_bytes(deflate_sync(&input, level))
        }),
    );

    vm.register_host_fn(
        "node:zlib",
        "inflateSync",
        Box::new(|_ctx, args| {
            let input = bytes_from_value(args.first().unwrap_or(&Value::Undefined));
            buf_from_bytes(inflate_sync(&input))
        }),
    );

    vm.register_host_fn(
        "node:zlib",
        "gzipSync",
        Box::new(|_ctx, args| {
            let input = bytes_from_value(args.first().unwrap_or(&Value::Undefined));
            let level = compression_level(args.get(1));
            buf_from_bytes(gzip_sync(&input, level))
        }),
    );

    vm.register_host_fn(
        "node:zlib",
        "gunzipSync",
        Box::new(|_ctx, args| {
            let input = bytes_from_value(args.first().unwrap_or(&Value::Undefined));
            buf_from_bytes(gunzip_sync(&input))
        }),
    );

    vm.register_host_fn(
        "node:zlib",
        "deflateRawSync",
        Box::new(|_ctx, args| {
            let input = bytes_from_value(args.first().unwrap_or(&Value::Undefined));
            let level = compression_level(args.get(1));
            buf_from_bytes(deflate_raw_sync(&input, level))
        }),
    );

    vm.register_host_fn(
        "node:zlib",
        "inflateRawSync",
        Box::new(|_ctx, args| {
            let input = bytes_from_value(args.first().unwrap_or(&Value::Undefined));
            buf_from_bytes(inflate_raw_sync(&input))
        }),
    );

    vm.register_host_fn(
        "node:zlib",
        "brotliCompressSync",
        Box::new(|_ctx, args| {
            let input = bytes_from_value(args.first().unwrap_or(&Value::Undefined));
            buf_from_bytes(brotli_compress_sync(&input))
        }),
    );

    vm.register_host_fn(
        "node:zlib",
        "brotliDecompressSync",
        Box::new(|_ctx, args| {
            let input = bytes_from_value(args.first().unwrap_or(&Value::Undefined));
            buf_from_bytes(brotli_decompress_sync(&input))
        }),
    );

    vm.register_host_fn(
        "node:zlib",
        "unzipSync",
        Box::new(|_ctx, args| {
            let input = bytes_from_value(args.first().unwrap_or(&Value::Undefined));
            buf_from_bytes(unzip_sync(&input))
        }),
    );

    vm.register_host_fn(
        "node:zlib",
        "constants",
        Box::new(|_ctx, _args| {
            let mut o = Object::new();
            // Return codes
            o.properties.insert("Z_OK".into(), Value::I32(0));
            o.properties.insert("Z_STREAM_END".into(), Value::I32(1));
            o.properties.insert("Z_NEED_DICT".into(), Value::I32(2));
            o.properties.insert("Z_ERRNO".into(), Value::I32(-1));
            o.properties.insert("Z_STREAM_ERROR".into(), Value::I32(-2));
            o.properties.insert("Z_DATA_ERROR".into(), Value::I32(-3));
            o.properties.insert("Z_MEM_ERROR".into(), Value::I32(-4));
            o.properties.insert("Z_BUF_ERROR".into(), Value::I32(-5));
            o.properties
                .insert("Z_VERSION_ERROR".into(), Value::I32(-6));
            // Compression levels
            o.properties
                .insert("Z_NO_COMPRESSION".into(), Value::I32(0));
            o.properties.insert("Z_BEST_SPEED".into(), Value::I32(1));
            o.properties
                .insert("Z_BEST_COMPRESSION".into(), Value::I32(9));
            o.properties
                .insert("Z_DEFAULT_COMPRESSION".into(), Value::I32(-1));
            // Strategies
            o.properties.insert("Z_FILTERED".into(), Value::I32(1));
            o.properties.insert("Z_HUFFMAN_ONLY".into(), Value::I32(2));
            o.properties.insert("Z_RLE".into(), Value::I32(3));
            o.properties.insert("Z_FIXED".into(), Value::I32(4));
            o.properties
                .insert("Z_DEFAULT_STRATEGY".into(), Value::I32(0));
            // Data types
            o.properties.insert("Z_BINARY".into(), Value::I32(0));
            o.properties.insert("Z_TEXT".into(), Value::I32(1));
            o.properties.insert("Z_ASCII".into(), Value::I32(1));
            o.properties.insert("Z_UNKNOWN".into(), Value::I32(2));
            // Flush constants
            o.properties.insert("Z_NO_FLUSH".into(), Value::I32(0));
            o.properties.insert("Z_PARTIAL_FLUSH".into(), Value::I32(1));
            o.properties.insert("Z_SYNC_FLUSH".into(), Value::I32(2));
            o.properties.insert("Z_FULL_FLUSH".into(), Value::I32(3));
            o.properties.insert("Z_FINISH".into(), Value::I32(4));
            o.properties.insert("Z_BLOCK".into(), Value::I32(5));
            o.properties.insert("Z_TREES".into(), Value::I32(6));
            // Brotli constants
            o.properties
                .insert("BROTLI_OPERATION_PROCESS".into(), Value::I32(0));
            o.properties
                .insert("BROTLI_OPERATION_FLUSH".into(), Value::I32(1));
            o.properties
                .insert("BROTLI_OPERATION_FINISH".into(), Value::I32(2));
            o.properties
                .insert("BROTLI_OPERATION_EMIT_METADATA".into(), Value::I32(3));
            o.properties
                .insert("BROTLI_PARAM_MODE".into(), Value::I32(0));
            o.properties
                .insert("BROTLI_PARAM_QUALITY".into(), Value::I32(1));
            o.properties
                .insert("BROTLI_PARAM_LGWIN".into(), Value::I32(2));
            Value::Object(vybe_runtime::heap::alloc(o))
        }),
    );

    // Async stubs (require event loop)
    for name in [
        "deflate",
        "inflate",
        "gzip",
        "gunzip",
        "deflateRaw",
        "inflateRaw",
        "brotliCompress",
        "brotliDecompress",
        "unzip",
    ] {
        vm.register_host_fn("node:zlib", name, Box::new(|_ctx, _args| Value::Undefined));
    }

    // Stream constructor stubs
    for name in [
        "createDeflate",
        "createInflate",
        "createGzip",
        "createGunzip",
        "createDeflateRaw",
        "createInflateRaw",
        "createBrotliCompress",
        "createBrotliDecompress",
        "createUnzip",
    ] {
        vm.register_host_fn("node:zlib", name, Box::new(|_ctx, _args| stub_stream()));
    }
}

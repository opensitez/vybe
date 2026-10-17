mod encoding {
    pub use vybe_platform_wasm::encoding::*;
}
use encoding::*;
use vybe_platform_wasm::writer::builtins::js_string_builtins as proposed;
fn section(out: &mut Vec<u8>, id: u8, bytes: &[u8]) {
    out.push(id);
    write_leb128_u32(out, bytes.len() as u32);
    out.extend(bytes);
}
fn name(out: &mut Vec<u8>, s: &str) {
    write_leb128_u32(out, s.len() as u32);
    out.extend(s.as_bytes());
}
fn module() -> Vec<u8> {
    let names = &proposed::IMPORTS[..13];
    let mut out = b"\0asm\x01\0\0\0".to_vec();
    let mut types = Vec::new();
    write_leb128_u32(&mut types, 17);
    types.extend([0x5e, 0x77, 1]);
    for n in names {
        types.push(TYPE_FUNC);
        assert!(proposed::write_signature(&mut types, n, 0));
    }
    types.extend([TYPE_FUNC, 1, TYPE_I32, 1, TYPE_EXTERNREF]);
    types.extend([TYPE_FUNC, 3, TYPE_EXTERNREF, TYPE_I32, TYPE_I32, 0]);
    types.extend([TYPE_FUNC, 2, TYPE_EXTERNREF, TYPE_I32, 1, TYPE_I32]);
    section(&mut out, 1, &types);
    let mut imports = vec![13];
    for (i, n) in names.iter().enumerate() {
        name(&mut imports, "wasm:js-string");
        name(&mut imports, n);
        imports.push(0);
        write_leb128_u32(&mut imports, i as u32 + 1);
    }
    section(&mut out, 2, &imports);
    section(&mut out, 3, &[3, 14, 15, 16]);
    let mut exports = vec![16];
    for (i, n) in names
        .iter()
        .chain(["gc_new", "gc_set", "gc_get"].iter())
        .enumerate()
    {
        name(&mut exports, n);
        exports.push(0);
        write_leb128_u32(&mut exports, i as u32);
    }
    section(&mut out, 7, &exports);
    let mut code = vec![3];
    for body in [
        vec![0, 0x20, 0, 0xfb, 7, 0, 0xfb, 0x1b, 0x0b],
        vec![
            0, 0x20, 0, 0xfb, 0x1a, 0xfb, 0x16, 0, 0x20, 1, 0x20, 2, 0xfb, 14, 0, 0x0b,
        ],
        vec![
            0, 0x20, 0, 0xfb, 0x1a, 0xfb, 0x16, 0, 0x20, 1, 0xfb, 13, 0, 0x0b,
        ],
    ] {
        write_leb128_u32(&mut code, body.len() as u32);
        code.extend(body);
    }
    section(&mut out, 10, &code);
    out
}
#[test]
fn all_standard_declarations_and_real_gc_accessors_validate() {
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&module())
        .unwrap();
}
#[test]
fn signed_heaptype_declarations_cover_high_concrete_ids() {
    for t in [0, 63, 64, 93, 127, 128, 65535] {
        let mut types = Vec::new();
        write_leb128_u32(&mut types, t + 2);
        for _ in 0..=t {
            types.extend([0x5e, 0x77, 1]);
        }
        types.push(TYPE_FUNC);
        assert!(proposed::write_signature(
            &mut types,
            "fromCharCodeArray",
            t
        ));
        let mut m = b"\0asm\x01\0\0\0".to_vec();
        section(&mut m, 1, &types);
        wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
            .validate_all(&m)
            .unwrap();
    }
}
#[test]
#[ignore = "root exports actual proposed declaration fixture for genuine V8 builtins"]
fn export_all_standard_declarations() {
    let m = module();
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&m)
        .unwrap();
    let output = std::path::PathBuf::from(
        std::env::var("JS_STRING_SIGNATURE_OUTPUT_DIR")
            .expect("explicit native export output directory"),
    );
    std::fs::create_dir_all(&output).unwrap();
    std::fs::write(output.join("standard-declarations.wasm"), m).unwrap();
}

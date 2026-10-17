//! Regression gates for module-owned memory and table declarations.

use crate::encoding::*;

fn module() -> Vec<u8> {
    let mut bytes = Vec::from(WASM_MAGIC);
    bytes.extend_from_slice(&WASM_VERSION);
    write_section(
        &mut bytes,
        SECTION_TYPE,
        &[2, TYPE_FUNC, 0, 1, TYPE_I32, TYPE_FUNC, 0, 0],
    );
    write_section(&mut bytes, SECTION_FUNCTION, &[2, 0, 1]);
    write_section(&mut bytes, 4, &[1, TYPE_FUNCREF, 1, 1, 2]);
    write_section(&mut bytes, SECTION_MEMORY, &[1, 1, 1, 2]);
    let mut exports = vec![1];
    write_name(&mut exports, "_start");
    exports.extend_from_slice(&[0, 1]);
    write_section(&mut bytes, SECTION_EXPORT, &exports);
    write_section(&mut bytes, 9, &[1, 0, 0x41, 0, 0x0b, 1, 0]);
    write_section(
        &mut bytes,
        SECTION_CODE,
        &[
            2, 4, 0, 0x41, 42, 0x0b, 17, 0, 0x41, 0, 0x28, 2, 0, 0x41, 0, 0x11, 0, 0, 0x47, 0x04,
            0x40, 0, 0x0b, 0x0b,
        ],
    );
    write_section(
        &mut bytes,
        SECTION_DATA,
        &[1, 0, 0x41, 0, 0x0b, 4, 42, 0, 0, 0],
    );
    bytes
}

fn validate(bytes: &[u8]) {
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(bytes)
        .expect("standard declaration fixture validates");
}

#[test]
fn native_declarations_are_retained_once_and_both_loaders_initialize_them() {
    let bytes = module();
    validate(&bytes);
    let chunks = super::read_wasm(&bytes).unwrap();
    assert_eq!(
        chunks[2].call_indirect_sigs.values().collect::<Vec<_>>(),
        vec!["->i32"]
    );
    assert_eq!(chunks[0].memory_min_pages, vec![1]);
    assert_eq!(chunks[0].memory_max_pages, vec![Some(2)]);
    assert_eq!(chunks[0].table_min_sizes, vec![1]);
    assert_eq!(chunks[0].table_max_sizes, vec![Some(2)]);
    for function in &chunks[1..] {
        assert!(function.memory_min_pages.is_empty());
        assert!(function.memory_max_pages.is_empty());
        assert!(function.table_min_sizes.is_empty());
        assert!(function.table_max_sizes.is_empty());
        assert!(function.table_element_types.is_empty());
        assert!(function.table_is_shared.is_empty());
        assert!(function.table_is_64.is_empty());
    }
    vybe_runtime::VM::new().run(chunks.clone()).unwrap();
    vybe_runtime::VM::new().run_linked(chunks, vec![]).unwrap();
}

#[test]
fn native_module_declarations_survive_standard_writer_roundtrip() {
    let bytes = module();
    validate(&bytes);
    let chunks = super::read_wasm(&bytes).unwrap();
    let written = crate::write_wasm(&chunks);
    validate(&written);
    let roundtrip = super::read_wasm(&written).unwrap();
    for linked in [false, true] {
        let mut vm = vybe_runtime::VM::new();
        vybe_runtime::js_builtins::register(&mut vm);
        vybe_platform_ecma::register(&mut vm);
        if linked {
            vm.run_linked(roundtrip.clone(), vec![]).unwrap();
        } else {
            vm.run(roundtrip.clone()).unwrap();
        }
    }
}

use crate::encoding::*;
use vybe_runtime::chunk::{ConstExpr, NativeSegmentMode};
use vybe_runtime::{VM, Value};

fn module() -> Vec<u8> {
    let mut bytes = Vec::from(WASM_MAGIC);
    bytes.extend_from_slice(&WASM_VERSION);
    bytes
}

fn decode(bytes: &[u8]) -> Vec<vybe_runtime::Chunk> {
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(bytes)
        .expect("standard native segment fixture");
    super::read_wasm(bytes).unwrap()
}

#[test]
fn imported_global_offsets_and_unsigned_bounds_work_on_both_loaders() {
    let mut bytes = module();
    let mut imports = vec![1];
    write_name(&mut imports, "env");
    write_name(&mut imports, "offset");
    imports.extend_from_slice(&[3, TYPE_I32, 0]);
    write_section(&mut bytes, SECTION_IMPORT, &imports);
    write_section(&mut bytes, SECTION_MEMORY, &[1, 0, 1]);
    write_section(
        &mut bytes,
        SECTION_DATA,
        &[1, 0, 0x23, 0, 0x0b, 3, 65, 66, 67],
    );
    let chunks = decode(&bytes);
    let metadata = chunks[0].native_segments.as_ref().unwrap();
    assert!(matches!(
        &metadata.data_modes[0],
        NativeSegmentMode::Active {
            index: 0,
            offset: ConstExpr::GlobalGet(_),
        }
    ));
    for linked in [false, true] {
        for offset in [7, -1] {
            let mut vm = VM::new();
            vm.set_global(
                &vybe_runtime::chunk::imported_global_key("env", "offset"),
                Value::I32(offset),
            );
            let result = if linked {
                vm.run_linked(chunks.clone(), vec![])
            } else {
                vm.run(chunks.clone())
            };
            if offset < 0 {
                assert!(result.is_err());
            } else {
                result.unwrap();
                vm.memory
                    .with_buffer_mut(|memory| assert_eq!(&memory[7..10], b"ABC"));
            }
        }
    }
}

#[test]
fn element_modes_reference_types_and_imported_functions_are_retained() {
    let mut bytes = module();
    write_section(
        &mut bytes,
        SECTION_TYPE,
        &[2, TYPE_FUNC, 0, 1, TYPE_I32, TYPE_FUNC, 0, 0],
    );
    let mut imports = vec![1];
    write_name(&mut imports, "env");
    write_name(&mut imports, "value");
    imports.extend_from_slice(&[0, 0]);
    write_section(&mut bytes, SECTION_IMPORT, &imports);
    write_section(&mut bytes, SECTION_FUNCTION, &[1, 1]);
    write_section(&mut bytes, 4, &[1, TYPE_FUNCREF, 0, 1]);
    let mut exports = vec![1];
    write_name(&mut exports, "_start");
    exports.extend_from_slice(&[0, 1]);
    write_section(&mut bytes, SECTION_EXPORT, &exports);
    write_section(
        &mut bytes,
        9,
        &[
            3, 0, 0x41, 0, 0x0b, 1, 0, 1, 0, 1, 0, 7, 0x6e, 1, 0xd0, 0x6e, 0x0b,
        ],
    );
    write_section(
        &mut bytes,
        SECTION_CODE,
        &[
            1, 14, 0, 0x41, 0, 0x11, 0, 0, 0x41, 42, 0x47, 0x04, 0x40, 0, 0x0b, 0x0b,
        ],
    );
    let chunks = decode(&bytes);
    let metadata = chunks[0].native_segments.as_ref().unwrap();
    assert_eq!(metadata.function_import_types, [0]);
    let segments = &metadata.elements;
    assert!(matches!(
        segments[0].mode,
        NativeSegmentMode::Active { index: 0, .. }
    ));
    assert!(matches!(segments[1].mode, NativeSegmentMode::Passive));
    assert!(matches!(segments[2].mode, NativeSegmentMode::Declarative));
    assert!(matches!(
        segments[0].items[0],
        ConstExpr::ImportedRefFunc(0)
    ));
    assert!(matches!(
        segments[1].items[0],
        ConstExpr::ImportedRefFunc(0)
    ));
    assert_eq!(
        segments[2].reference_type.0,
        vybe_runtime::opcode::heaptype::HeapType::Abstract(0x6e)
    );
    for linked in [false, true] {
        let mut vm = VM::new();
        let calls = std::sync::Arc::new(std::sync::atomic::AtomicUsize::new(0));
        let observed = calls.clone();
        vm.register_host_fn(
            "env",
            "value",
            Box::new(move |_, _| {
                observed.fetch_add(1, std::sync::atomic::Ordering::SeqCst);
                Value::I32(42)
            }),
        );
        if linked {
            let import = vm.resolve_import_target("env", "value").unwrap();
            vm.run_linked(chunks.clone(), vec![import]).unwrap();
        } else {
            vm.run(chunks.clone()).unwrap();
        }
        assert_eq!(calls.load(std::sync::atomic::Ordering::SeqCst), 1);
    }
}

#[test]
fn memory64_offsets_keep_full_width_and_trap_instead_of_wrapping() {
    let mut bytes = module();
    write_section(&mut bytes, SECTION_MEMORY, &[1, 4, 1]);
    write_section(&mut bytes, SECTION_DATA, &[1, 0, 0x42, 0x7f, 0x0b, 1, 42]);
    let chunks = decode(&bytes);
    assert!(matches!(
        &chunks[0].native_segments.as_ref().unwrap().data_modes[0],
        NativeSegmentMode::Active {
            offset: ConstExpr::Value(Value::I64(-1)),
            ..
        }
    ));
    assert!(VM::new().run(chunks.clone()).is_err());
    assert!(VM::new().run_linked(chunks, vec![]).is_err());
}

#[test]
fn active_data_targets_the_declared_nonzero_memory() {
    let mut bytes = module();
    write_section(&mut bytes, SECTION_TYPE, &[1, TYPE_FUNC, 0, 0]);
    write_section(&mut bytes, SECTION_FUNCTION, &[1, 0]);
    write_section(&mut bytes, SECTION_MEMORY, &[2, 0, 1, 4, 1]);
    let mut exports = vec![1];
    write_name(&mut exports, "_start");
    exports.extend_from_slice(&[0, 0]);
    write_section(&mut bytes, SECTION_EXPORT, &exports);
    let body = [
        0, 0x42, 2, 0x28, 0x42, 1, 0, 0x41, 42, 0x47, 0x04, 0x40, 0, 0x0b, 0x0b,
    ];
    let mut code = vec![1, body.len() as u8];
    code.extend_from_slice(&body);
    write_section(&mut bytes, SECTION_CODE, &code);
    write_section(&mut bytes, SECTION_DATA, &[1, 2, 1, 0x42, 2, 0x0b, 1, 42]);
    let chunks = decode(&bytes);
    assert!(matches!(
        &chunks[0].native_segments.as_ref().unwrap().data_modes[0],
        NativeSegmentMode::Active {
            index: 1,
            offset: ConstExpr::Value(Value::I64(2))
        }
    ));
    for linked in [false, true] {
        let mut vm = VM::new();
        if linked {
            vm.run_linked(chunks.clone(), vec![]).unwrap();
        } else {
            vm.run(chunks.clone()).unwrap();
        }
        vm.memory.with_buffer_mut(|memory| assert_eq!(memory[2], 0));
    }
}

#[test]
fn native_global_declarations_preserve_types_mutability_and_sharedness() {
    use vybe_runtime::chunk::NativeValueType as Type;
    use vybe_runtime::opcode::heaptype::HeapType;
    let mut bytes = module();
    write_section(&mut bytes, SECTION_TYPE, &[1, 0x5f, 0]);
    let mut imports = vec![1];
    write_name(&mut imports, "env");
    write_name(&mut imports, "seed");
    imports.extend_from_slice(&[3, TYPE_I64, 0]);
    write_section(&mut bytes, SECTION_IMPORT, &imports);
    let mut globals = vec![7, TYPE_I32, 1, 0x41, 7, 0x0b, TYPE_F32, 0, 0x43];
    globals.extend_from_slice(&1.5f32.to_le_bytes());
    globals.extend_from_slice(&[0x0b, TYPE_F64, 1, 0x44]);
    globals.extend_from_slice(&2.5f64.to_le_bytes());
    globals.extend_from_slice(&[0x0b, 0x7b, 0, 0xfd, 12]);
    globals.extend_from_slice(&[0; 16]);
    globals.extend_from_slice(&[
        0x0b, 0x6e, 0, 0xd0, 0x6e, 0x0b, TYPE_I64, 3, 0x42, 1, 0x0b, 0x63, 0, 0, 0xd0, 0, 0x0b,
    ]);
    write_section(&mut bytes, SECTION_GLOBAL, &globals);
    let chunks = decode(&bytes);
    let declarations = &chunks[0].native_segments.as_ref().unwrap().global_types;
    assert_eq!(declarations.len(), 8);
    assert_eq!(
        declarations
            .iter()
            .map(|global| global.value_type.clone())
            .collect::<Vec<_>>(),
        vec![
            Type::I64,
            Type::I32,
            Type::F32,
            Type::F64,
            Type::V128,
            Type::Ref((HeapType::Abstract(0x6e), true, false, false)),
            Type::I64,
            Type::Ref((HeapType::Concrete(1), true, false, false))
        ]
    );
    assert_eq!(
        declarations
            .iter()
            .map(|global| global.mutable)
            .collect::<Vec<_>>(),
        vec![false, true, false, true, false, false, true, false]
    );
    assert_eq!(
        declarations
            .iter()
            .map(|global| global.shared)
            .collect::<Vec<_>>(),
        vec![false, false, false, false, false, false, true, false]
    );
    assert!(
        chunks
            .iter()
            .skip(1)
            .all(|chunk| chunk.native_segments.is_none())
    );
}

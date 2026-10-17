use crate::encoding::*;

fn module(tail: bool, mismatch: bool) -> Vec<u8> {
    let mut bytes = Vec::from(WASM_MAGIC);
    bytes.extend_from_slice(&WASM_VERSION);
    write_section(
        &mut bytes,
        SECTION_TYPE,
        &[2, TYPE_FUNC, 0, 1, TYPE_I32, TYPE_FUNC, 0, 1, TYPE_I64],
    );
    write_section(&mut bytes, SECTION_FUNCTION, &[2, u8::from(mismatch), 0]);
    write_section(&mut bytes, 4, &[1, TYPE_FUNCREF, 0, 1]);
    let mut exports = vec![1];
    write_name(&mut exports, "_start");
    exports.extend_from_slice(&[0, 1]);
    write_section(&mut bytes, SECTION_EXPORT, &exports);
    write_section(&mut bytes, 9, &[1, 0, 0x41, 0, 0x0b, 1, 0]);
    write_section(
        &mut bytes,
        SECTION_CODE,
        &[
            2,
            4,
            0,
            if mismatch { 0x42 } else { 0x41 },
            42,
            0x0b,
            7,
            0,
            0x41,
            0,
            if tail { 0x13 } else { 0x11 },
            0,
            0,
            0x0b,
        ],
    );
    bytes
}

fn read(tail: bool, mismatch: bool) -> Vec<vybe_runtime::Chunk> {
    let bytes = module(tail, mismatch);
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&bytes)
        .expect("indirect target mismatch is a runtime check");
    super::read_wasm(&bytes).unwrap()
}

#[test]
fn normal_and_tail_indirect_calls_retain_declared_signatures() {
    for tail in [false, true] {
        let chunks = read(tail, false);
        let signatures: Vec<_> = chunks[2]
            .call_indirect_sigs
            .values()
            .map(String::as_str)
            .collect();
        assert_eq!(signatures, vec!["->i32"]);
        vybe_runtime::VM::new().run(chunks.clone()).unwrap();
        vybe_runtime::VM::new().run_linked(chunks, vec![]).unwrap();
    }
}

#[test]
fn same_arity_different_result_types_trap_on_both_loaders() {
    for tail in [false, true] {
        let chunks = read(tail, true);
        for linked in [false, true] {
            let mut vm = vybe_runtime::VM::new();
            let result = if linked {
                vm.run_linked(chunks.clone(), vec![])
            } else {
                vm.run(chunks.clone())
            };
            let error = result.expect_err("i64 target cannot satisfy an i32 call declaration");
            assert!(error.is_trap(), "{error:?}");
            assert!(
                error.message.contains("indirect call type mismatch"),
                "{error:?}"
            );
        }
    }
}

// Actual decoded standard Wasm; Source carrier and VM opcode layout unchanged.
#[test]
fn actual_declared_reference_calls_typed_arguments_and_returns() {
    let dir=std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/declared_ref_calls");
    for name in ["i32","i64","f32","f64","externref","funcref","mixed","tail_i32","multiple"] {
        let original=std::fs::read(dir.join(format!("{name}.wasm"))).unwrap();
        wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all()).validate_all(&original).unwrap();
        let chunks=vybe_platform_wasm::read_wasm(&original).unwrap();
        for c in &chunks { println!("{name} {} sig{:?} indirect{:?} locals{:?} code{:?}",c.name,c.func_sig,c.call_indirect_sigs,c.local_names,c.code); }
        let rewritten=vybe_platform_wasm::write_wasm(&chunks);
        
        let result=wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all()).validate_all(&rewritten);
        assert!(result.is_ok(),"{name}: {:?}",result.err());
    }
}
#[test]
fn reader_signature_metadata_preserves_actual_vm_values() {
    use vybe_runtime::{Value,VM,Op};
    let dir=std::path::Path::new(env!("CARGO_MANIFEST_DIR")).join("tests/fixtures/declared_ref_calls");
    for (name,arg) in [("i32",Value::I32(-17)),("i64",Value::I64(-17)),("f32",Value::F32(1.5)),("f64",Value::F64(-0.)),("externref",Value::String("x".into()))] {
        let mut chunks=vybe_platform_wasm::read_wasm(&std::fs::read(dir.join(format!("{name}.wasm"))).unwrap()).unwrap();
        // VM host-entry adapter alone: original callee bodies/imports/metadata
        // unchanged; native gates use the original full modules above.
        chunks[0].code.clear();chunks[0].emit_op_u16(Op::REF_FUNC,2,0);chunks[0].emit(0,0);chunks[0].emit_op(Op::RETURN,0);
        let mut vm=VM::new();let callback=vm.run(chunks).unwrap();
        assert_eq!(vm.invoke(&callback,&[arg.clone()]).unwrap(),arg,"{name}");
    }
}
#[test]
fn source_spelling_cannot_change_declared_native_local_types() {
    use vybe_runtime::chunk::{LocalName,LocalOrigin};
    let original=include_bytes!("fixtures/declared_ref_calls/i32.wasm");
    let chunks=vybe_platform_wasm::read_wasm(original).unwrap();
    let expected=vybe_platform_wasm::write_wasm(&chunks);
    let mut spoofed=chunks.clone();
    spoofed[2].local_names.push(LocalName{slot:1,name:"<native-wasm-scratch:funcref>".into(),origin:LocalOrigin::Source});
    assert_eq!(vybe_platform_wasm::write_wasm(&spoofed),expected);
}

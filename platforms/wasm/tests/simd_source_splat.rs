use vybe_runtime::{Chunk, Op, Value, VM};
use vybe_runtime::chunk::{LocalName, LocalOrigin};

fn splat(declared: bool, scratch: bool) -> Chunk {
    let mut c = Chunk::new("splat");
    c.arity = 1; c.param_count = 1; c.local_count = 1;
    if declared { c.func_sig = Some(("i32".into(), "i32".into())); }
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_op(Op::I8X16_SPLAT, 0);
    if scratch {
        c.local_count=2;
        c.local_names.push(LocalName::new(1,"<native-wasm-scratch:v128>",LocalOrigin::Compiler));
        c.emit_op_u16(Op::LOCAL_SET,1,0);
        c.emit_op_u16(Op::LOCAL_GET,1,0);
    }
    c.emit_op(Op::I8X16_BITMASK, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

#[test]
fn byte_splat_accepts_source_scalars_and_declared_i32_parameters() {
    for declared in [false, true] {for scratch in [false,true] {
        let c = splat(declared,scratch);
        wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
            .validate_all(&vybe_platform_wasm::write_wasm(&[c])).unwrap();
    }}
    let mut root = Chunk::new("owner");
    root.emit_op_u16(Op::REF_FUNC, 1, 0); root.emit(0, 0); root.emit_op(Op::RETURN, 0);
    let mut vm = VM::new();
    let function = vm.run(vec![root, splat(false,true)]).unwrap();
    for value in [0, 1, 127, 128, 255, 511, -1] {
        let expected = if value & 128 != 0 { 65535.0 } else { 0.0 };
        assert_eq!(vm.invoke(&function, &[Value::I32(value)]).unwrap().as_f64(), expected);
    }
}

fn byte_shift(declared:bool)->Chunk {
    let mut c=Chunk::new("byte_shift");c.arity=1;c.param_count=1;c.local_count=1;
    if declared {c.func_sig=Some(("i32".into(),"i32".into()));}
    c.emit_op(Op::V128_CONST,0);for _ in 0..16 {c.emit(128,0);}
    c.emit_op_u16(Op::LOCAL_GET,0,0);c.emit_op(Op::I8X16_SHR_U,0);c.emit_op(Op::I8X16_BITMASK,0);c.emit_op(Op::RETURN,0);c
}
#[test]
fn byte_shift_unboxes_source_scalar_counts() {
    for declared in [false,true] {wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all()).validate_all(&vybe_platform_wasm::write_wasm(&[byte_shift(declared)])).unwrap();}
    let mut root=Chunk::new("owner");root.emit_op_u16(Op::REF_FUNC,1,0);root.emit(0,0);root.emit_op(Op::RETURN,0);
    let mut vm=VM::new();let function=vm.run(vec![root,byte_shift(false)]).unwrap();
    for count in [0,1,4,7,8,9,32,-1,i32::MIN,i32::MAX] {let expected=if count&7==0 {65535.0}else{0.0};assert_eq!(vm.invoke(&function,&[Value::I32(count)]).unwrap().as_f64(),expected);}
}

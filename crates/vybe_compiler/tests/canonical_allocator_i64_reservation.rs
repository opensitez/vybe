use vybe_compiler::primitives;
use vybe_runtime::{Chunk, Op, Value, VM};
fn module() -> Vec<Chunk> {
    let mut c = vec![Chunk::new("allocator-owner"), Chunk::new("allocator-call")];
    c[0].memory_min_pages.push(1);
    c[1].arity = 1; c[1].param_count = 1; c[1].local_count = 1;
    // An existing raw lane from an earlier completed canonical/bitwise use.
    let reserved = c[1].i64_scratch();
    c[1].emit_i64_const(7, 0);
    c[1].emit_op_u16(Op::LOCAL_SET, reserved, 0);
    let ptr = primitives::canon_marshal::emit_alloc(&mut c[1], 0, 0);
    assert_eq!(c[1].i64_scratch_slot, Some(reserved));
    assert_eq!(c[1].i64_locals.len(), 1, "reuse one reserved i64 lane");
    let mut at = 0; let mut seen = false;
    while at < c[1].code.len() {
        let op = Op::from_code_at(&c[1].code, at);
        if op == Op::LOCAL_TEE && at >= 4 && c[1].read_u16(at+4) == reserved { seen = true; }
        at += 4 + op.operand_format().size_in(&c[1].code, at+4);
    }
    assert!(seen, "allocator stores checked widened next in reserved lane");
    c[1].emit_op_u16(Op::LOCAL_GET, ptr, 0); c[1].emit_op(Op::RETURN, 0);
    primitives::functions::emit_ref_func(&mut c[0], 1, 0, 0);
    c[0].emit_op(Op::RETURN, 0);
    primitives::globals::declare_free_globals(&mut c);
    primitives::globals::normalize_global_table(&mut c); c
}
#[test]
fn reserved_allocator_lane_complete_native_module() {
    let c=module(); let bytes=vybe_platform_wasm::write_wasm(&c);
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all()).validate_all(&bytes).unwrap();
    let mut seen=false;
    for payload in wasmparser::Parser::new(0).parse_all(&bytes) {
        if let wasmparser::Payload::CodeSectionEntry(body)=payload.unwrap() {
            for row in body.get_locals_reader().unwrap() {let (_,ty)=row.unwrap(); seen |= ty==wasmparser::ValType::I64;}
        }
    }
    assert!(seen,"native raw i64 local declaration retained");
}

#[test]
fn allocator_raw_lane_does_not_retype_reclaimed_value_slots() {
    let mut chunks = module();
    let chunk = &mut chunks[1];
    chunk.code.clear();
    chunk.i64_locals.clear();
    chunk.i64_scratch_slot = None;
    chunk.local_count = 1;
    let value = chunk.alloc_scratch(8);
    chunk.emit_i32_const(42, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, value, 0);
    let prior_peak = chunk.local_count;
    chunk.local_count = 1;
    let ptr = primitives::canon_marshal::emit_alloc(chunk, 0, 0);
    assert!(chunk.i64_scratch_slot.unwrap() >= prior_peak);
    chunk.emit_op_u16(Op::LOCAL_GET, ptr, 0);
    chunk.emit_op(Op::RETURN, 0);
    let bytes = vybe_platform_wasm::write_wasm(&chunks);
    wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
        .validate_all(&bytes).unwrap();
}
#[test]
fn empty_small_and_multi_page_allocations_preserve_bump_and_alignment() {
    let mut vm=VM::new();primitives::platforms::register_platforms_all(&mut vm);
    let f=vm.run(module()).unwrap();
    let mut expected=65536u32;
    for size in [0u32,17,256*1024+17,0] {
        assert_eq!(vm.invoke(&f,&[Value::I32(size as i32)]).unwrap(),Value::I32(expected as i32));
        expected += (size+15)&!7;
    }
}
#[test]
fn overflowing_memory32_request_traps_before_giant_growth() {
    let mut vm=VM::new();primitives::platforms::register_platforms_all(&mut vm);
    let f=vm.run(module()).unwrap();
    assert!(vm.invoke(&f,&[Value::I32(-1)]).is_err());
}

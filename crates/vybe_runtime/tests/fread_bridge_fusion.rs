//! Runtime coverage for the direct-store offset loop emitted by cached fread.
use std::time::Instant;
use vybe_runtime::{Chunk, Op, Value, VM};

#[test]
fn emitted_fread_offset_bridge_uses_bulk_fusion() {
    const SRC_OFFSET: usize = 7;
    const TOTAL: usize = 32 * 1024;
    const SOURCE_LEN: usize = SRC_OFFSET + TOTAL;
    const ARR_GLOBAL: &str = "__perf_managed_array_offset_source";

    const ARR_SLOT: u16 = 0;
    const DST_SLOT: u16 = 1;
    const LEN_SLOT: u16 = 2;
    const I_SLOT: u16 = 3;
    const BYTE_SLOT: u16 = 4;
    const SRC_OFFSET_SLOT: u16 = 7;

    let mut setup = Chunk::new("<managed-array-offset-source-setup>");
    for i in 0..SOURCE_LEN {
        setup.emit_i32_const((i & 0xff) as i32, 0);
    }
    setup.emit_array_new_fixed(0, SOURCE_LEN as u16, 0);
    setup.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    let source = vm.run(vec![setup]).expect("offset bridge source setup failed");
    vm.set_global(ARR_GLOBAL, source);

    let mut chunk = Chunk::new("<managed-array-offset-to-linear-bridge>");
    chunk.local_count = 8;

    chunk.emit_f64_const(1.0, 0);
    chunk.emit_op(Op::MEMORY_GROW, 0);
    chunk.emit_leb_u32(0, 0);
    chunk.emit_op(Op::DROP, 0);

    let arr_global_idx = chunk.add_constant(Value::String(ARR_GLOBAL.into()));
    chunk.emit_op_u32(Op::GLOBAL_GET, arr_global_idx, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, ARR_SLOT, 0);

    chunk.emit_i32_const(0, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, DST_SLOT, 0);
    chunk.emit_i32_const(TOTAL as i32, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, LEN_SLOT, 0);
    chunk.emit_i32_const(0, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, I_SLOT, 0);
    chunk.emit_i32_const(SRC_OFFSET as i32, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, SRC_OFFSET_SLOT, 0);

    let outer = chunk.emit_block(0);
    let (lp, _loop_start) = chunk.emit_loop_s(0);

    chunk.emit_op_u16(Op::LOCAL_GET, I_SLOT, 0);
    chunk.emit_op_u16(Op::LOCAL_GET, LEN_SLOT, 0);
    chunk.emit_op(Op::I32_GE_U, 0);
    chunk.emit_br_if(1, 0);

    chunk.emit_op_u16(Op::LOCAL_GET, ARR_SLOT, 0);
    chunk.emit_op_u16(Op::LOCAL_GET, SRC_OFFSET_SLOT, 0);
    chunk.emit_op_u16(Op::LOCAL_GET, I_SLOT, 0);
    chunk.emit_op(Op::I32_ADD, 0);
    chunk.emit_op(Op::ARRAY_GET, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, BYTE_SLOT, 0);

    chunk.emit_op_u16(Op::LOCAL_GET, DST_SLOT, 0);
    chunk.emit_op_u16(Op::LOCAL_GET, I_SLOT, 0);
    chunk.emit_op(Op::I32_ADD, 0);
    chunk.emit_op_u16(Op::LOCAL_GET, BYTE_SLOT, 0);
    chunk.emit_op(Op::I32_STORE8, 0);

    chunk.emit_op_u16(Op::LOCAL_GET, I_SLOT, 0);
    chunk.emit_i32_const(1, 0);
    chunk.emit_op(Op::I32_ADD, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, I_SLOT, 0);
    chunk.emit_br(0, 0);

    chunk.emit_end(0);
    chunk.patch_loop(lp);
    chunk.emit_end(0);
    chunk.patch_block(outer);

    chunk.emit_op_u16(Op::LOCAL_GET, DST_SLOT, 0);
    chunk.emit_op_u16(Op::LOCAL_GET, LEN_SLOT, 0);
    chunk.emit_i32_const(1, 0);
    chunk.emit_op(Op::I32_SUB, 0);
    chunk.emit_op(Op::I32_ADD, 0);
    chunk.emit_op(Op::I32_LOAD8_U, 0);
    chunk.emit_op(Op::RETURN, 0);

    vm.record_runtime_perf(true);
    let start = Instant::now();
    let result = vm.run(vec![chunk]).expect("offset bridge fixture failed");
    let elapsed = start.elapsed();
    assert_eq!(result, Value::I32(((SRC_OFFSET + TOTAL - 1) & 0xff) as i32));

    let counters = vm.take_runtime_perf().expect("perf counters enabled");
    println!(
        "managed array offset -> linear bridge: {} bytes in {:.2} ms ({:.1} ns/byte)",
        TOTAL,
        elapsed.as_secs_f64() * 1000.0,
        elapsed.as_nanos() as f64 / TOTAL as f64
    );
    println!("{}", counters.format_report());
    println!("{}", counters.bridge_summary());
    assert_eq!(counters.super_managed_to_linear, 1);
    assert_eq!(counters.array_get, TOTAL as u64);
    assert_eq!(counters.i32_store8, TOTAL as u64);
}


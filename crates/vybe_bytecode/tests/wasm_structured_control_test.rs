use vybe_bytecode::value::Value;
use vybe_bytecode::{Chunk, Op, VM};

fn run_chunk_expect_i32(chunk: Chunk, expected: i32) {
    match VM::new().run(vec![chunk]).expect("chunk should execute") {
        Value::I32(n) => assert_eq!(n, expected, "expected i32 {expected}"),
        other => panic!("expected i32 {expected}, got {other:?}"),
    }
}

fn make_i32(chunk: &mut Chunk, v: i32) {
    let idx = chunk.add_constant(Value::I32(v));
    chunk.emit_op_u16(Op::CONST, idx, 0);
}

fn emit_if_typed(chunk: &mut Chunk, result_count: u8) {
    chunk.emit_op(Op::IF, 0);
    chunk.emit(result_count, 0);
}

fn leb(out: &mut Vec<u8>, mut value: u32) {
    loop {
        let mut byte = (value & 0x7f) as u8;
        value >>= 7;
        if value != 0 {
            byte |= 0x80;
        }
        out.push(byte);
        if value == 0 {
            break;
        }
    }
}

fn section(out: &mut Vec<u8>, id: u8, payload: Vec<u8>) {
    out.push(id);
    leb(out, payload.len() as u32);
    out.extend_from_slice(&payload);
}

fn standard_wasm_function(instructions: &[u8]) -> Vec<u8> {
    let mut wasm = Vec::new();
    wasm.extend_from_slice(&[0x00, 0x61, 0x73, 0x6d]);
    wasm.extend_from_slice(&[0x01, 0x00, 0x00, 0x00]);

    let type_section = vec![0x01, 0x60, 0x00, 0x01, 0x7f];
    section(&mut wasm, 1, type_section);

    let function_section = vec![0x01, 0x00];
    section(&mut wasm, 3, function_section);

    let mut body = Vec::new();
    body.push(0x00); // local decl count
    body.extend_from_slice(instructions);
    body.push(0x0b); // function end

    let mut code_section = Vec::new();
    code_section.push(0x01);
    leb(&mut code_section, body.len() as u32);
    code_section.extend_from_slice(&body);
    section(&mut wasm, 10, code_section);

    wasm
}

fn standard_wasm_function_with_type(
    type_entries: &[u8],
    func_type_idx: u32,
    instructions: &[u8],
) -> Vec<u8> {
    let mut wasm = Vec::new();
    wasm.extend_from_slice(&[0x00, 0x61, 0x73, 0x6d]);
    wasm.extend_from_slice(&[0x01, 0x00, 0x00, 0x00]);

    section(&mut wasm, 1, type_entries.to_vec());

    let mut function_section = vec![0x01];
    leb(&mut function_section, func_type_idx);
    section(&mut wasm, 3, function_section);

    let mut body = Vec::new();
    body.push(0x00);
    body.extend_from_slice(instructions);
    body.push(0x0b);

    let mut code_section = Vec::new();
    code_section.push(0x01);
    leb(&mut code_section, body.len() as u32);
    code_section.extend_from_slice(&body);
    section(&mut wasm, 10, code_section);

    wasm
}

fn decoded_function(instructions: &[u8]) -> Chunk {
    let wasm = standard_wasm_function(instructions);
    let mut chunks = vybe_bytecode::wasm::read_wasm(&wasm).expect("standard wasm should read");
    assert_eq!(chunks.len(), 2);
    chunks.remove(1)
}

fn run_chunk(chunk: Chunk) -> Value {
    VM::new().run(vec![chunk]).expect("chunk should execute")
}

#[test]
fn reader_preserves_if_else_and_i32_conditions() {
    let chunk = decoded_function(&[
        0x41, 0x00, // i32.const 0
        0x04, 0x7f, // if (result i32)
        0x41, 0x0b, //   i32.const 11
        0x05, // else
        0x41, 0x16, //   i32.const 22
        0x0b, // end
    ]);

    assert!(
        chunk
            .code
            .windows(2)
            .any(|w| w == [Op::IF.prefix(), Op::IF.sub()])
    );
    assert!(
        chunk
            .code
            .windows(2)
            .any(|w| w == [Op::ELSE.prefix(), Op::ELSE.sub()])
    );
    assert!(
        chunk
            .code
            .windows(2)
            .any(|w| w == [Op::END.prefix(), Op::END.sub()])
    );
    assert_eq!(run_chunk(chunk).as_i32(), 22);
}

#[test]
fn multi_value_if_branch_preserves_results_and_drops_temps() {
    let mut chunk = Chunk::new("<script>");
    make_i32(&mut chunk, 99); // value below the if
    make_i32(&mut chunk, 1); // true condition
    emit_if_typed(&mut chunk, 2);
    make_i32(&mut chunk, 77); // branch-local temp, not a result
    make_i32(&mut chunk, 1); // first branch result
    make_i32(&mut chunk, 2); // second branch result
    chunk.emit_br(0, 0);
    chunk.emit_end(0);
    chunk.emit_op(Op::I32_ADD, 0); // 1 + 2
    chunk.emit_op(Op::I32_ADD, 0); // 99 + 3; would be 77 + 3 without stack shaping
    chunk.emit_op(Op::RETURN, 0);

    run_chunk_expect_i32(chunk, 102);
}

#[test]
fn reader_preserves_br_if_depth() {
    let chunk = decoded_function(&[
        0x02, 0x40, // block
        0x41, 0x01, //   i32.const 1
        0x0d, 0x00, //   br_if 0
        0x41, 0x07, //   i32.const 7
        0x0b, // end
        0x41, 0x09, // i32.const 9
    ]);

    assert!(
        chunk
            .code
            .windows(2)
            .any(|w| w == [Op::BR_IF.prefix(), Op::BR_IF.sub()])
    );
    assert_eq!(run_chunk(chunk).as_i32(), 9);
}

#[test]
fn reader_preserves_br_table_vector_and_default_depths() {
    let chunk = decoded_function(&[
        0x02, 0x40, // block outer
        0x02, 0x40, //   block inner
        0x41, 0x01, //     i32.const 1
        0x0e, 0x02, //     br_table count=2
        0x00, 0x01, 0x01, //       labels [0, 1], default 1
        0x41, 0x07, //     i32.const 7
        0x0b, //   end inner
        0x41, 0x08, //   i32.const 8
        0x0b, // end outer
        0x41, 0x09, // i32.const 9
    ]);

    let br_table = chunk
        .code
        .windows(2)
        .position(|w| w == [Op::BR_TABLE.prefix(), Op::BR_TABLE.sub()])
        .expect("br_table should be decoded");
    assert_eq!(
        &chunk.code[br_table + 2..br_table + 6],
        &[0x02, 0x00, 0x01, 0x01]
    );
    assert_eq!(run_chunk(chunk).as_i32(), 9);
}

#[test]
fn vm_rejects_non_numeric_if_condition() {
    let mut chunk = Chunk::new("<script>");
    chunk.emit_op(Op::NULL, 0);
    chunk.emit_if(0);
    chunk.emit_op(Op::NULL, 0);
    chunk.emit_op(Op::END, 0);
    chunk.emit_op(Op::RETURN, 0);

    let err = VM::new()
        .run(vec![chunk])
        .expect_err("if must require an i32 condition");
    assert!(err.message.contains("if expected i32 condition"));
}

#[test]
fn core_comparisons_produce_i32_zero_or_one() {
    let mut chunk = Chunk::new("<script>");
    let zero = chunk.add_constant(Value::I32(0));
    chunk.emit_op_u16(Op::CONST, zero, 0);
    chunk.emit_op(Op::I32_EQZ, 0);
    chunk.emit_op(Op::RETURN, 0);
    assert!(matches!(run_chunk(chunk), Value::I32(1)));

    let mut chunk = Chunk::new("<script>");
    let a = chunk.add_constant(Value::I32(7));
    let b = chunk.add_constant(Value::I32(7));
    chunk.emit_op_u16(Op::CONST, a, 0);
    chunk.emit_op_u16(Op::CONST, b, 0);
    chunk.emit_op(Op::I32_EQ, 0);
    chunk.emit_op(Op::RETURN, 0);
    assert!(matches!(run_chunk(chunk), Value::I32(1)));
}

#[test]
fn reader_uses_function_section_type_indices() {
    let types = [
        0x02, // two function types
        0x60, 0x00, 0x00, // type 0: () -> ()
        0x60, 0x02, 0x7f, 0x7f, 0x01, 0x7f, // type 1: (i32, i32) -> i32
    ];
    let wasm = standard_wasm_function_with_type(&types, 1, &[0x41, 0x00]);
    let chunks = vybe_bytecode::wasm::read_wasm(&wasm).expect("standard wasm should read");
    assert_eq!(chunks[1].arity, 2);
    assert_eq!(chunks[1].result_arity, 1);
}

#[test]
fn reader_decodes_prefixed_proposal_opcodes() {
    let chunk = decoded_function(&[
        0x41, 0x01, // i32.const 1
        0xfb, 0x1c, // ref.i31
        0x1a, // drop
        0xfd, 0x0c, // v128.const
        0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 0x1a, // drop
        0xfe, 0x03, 0x00, // atomic.fence 0
        0x41, 0x00, // i32.const 0
    ]);

    assert!(
        chunk
            .code
            .windows(2)
            .any(|w| w == [Op::I31_NEW.prefix(), Op::I31_NEW.sub()])
    );
    assert!(
        chunk
            .code
            .windows(2)
            .any(|w| w == [Op::V128_CONST.prefix(), Op::V128_CONST.sub()])
    );
    assert!(
        chunk
            .code
            .windows(2)
            .any(|w| w == [Op::ATOMIC_FENCE.prefix(), Op::ATOMIC_FENCE.sub()])
    );
}

// ── VM execution: if/else ─────────────────────────────────────────────────

#[test]
fn if_true_branch_executes() {
    // if (1) { return 42 } end; return 0
    let mut c = Chunk::new("<script>");
    make_i32(&mut c, 1);
    c.emit_if(0);
    make_i32(&mut c, 42);
    c.emit_op(Op::RETURN, 0);
    c.emit_end(0);
    make_i32(&mut c, 0);
    c.emit_op(Op::RETURN, 0);
    run_chunk_expect_i32(c, 42);
}

#[test]
fn if_false_skips_body() {
    // if (0) { return 99 } end; return 7
    let mut c = Chunk::new("<script>");
    make_i32(&mut c, 0);
    c.emit_if(0);
    make_i32(&mut c, 99);
    c.emit_op(Op::RETURN, 0);
    c.emit_end(0);
    make_i32(&mut c, 7);
    c.emit_op(Op::RETURN, 0);
    run_chunk_expect_i32(c, 7);
}

#[test]
fn if_false_skips_unreachable_body() {
    let mut c = Chunk::new("<script>");
    make_i32(&mut c, 0);
    c.emit_if(0);
    c.emit_op(Op::UNREACHABLE, 0);
    c.emit_end(0);
    make_i32(&mut c, 7);
    c.emit_op(Op::RETURN, 0);
    run_chunk_expect_i32(c, 7);
}

#[test]
fn if_else_takes_then_when_true() {
    // if (1) { 11 } else { 22 } end; return
    let mut c = Chunk::new("<script>");
    make_i32(&mut c, 1);
    c.emit_if_value(0);
    make_i32(&mut c, 11);
    c.emit_else(0);
    make_i32(&mut c, 22);
    c.emit_end(0);
    c.emit_op(Op::RETURN, 0);
    run_chunk_expect_i32(c, 11);
}

#[test]
fn if_else_takes_else_when_false() {
    // if (0) { 11 } else { 22 } end; return
    let mut c = Chunk::new("<script>");
    make_i32(&mut c, 0);
    c.emit_if_value(0);
    make_i32(&mut c, 11);
    c.emit_else(0);
    make_i32(&mut c, 22);
    c.emit_end(0);
    c.emit_op(Op::RETURN, 0);
    run_chunk_expect_i32(c, 22);
}

#[test]
fn nested_if_inner_true() {
    // if (1) { if (1) { return 5 } end; return 6 } end; return 7
    let mut c = Chunk::new("<script>");
    make_i32(&mut c, 1);
    c.emit_if(0);
    make_i32(&mut c, 1);
    c.emit_if(0);
    make_i32(&mut c, 5);
    c.emit_op(Op::RETURN, 0);
    c.emit_end(0);
    make_i32(&mut c, 6);
    c.emit_op(Op::RETURN, 0);
    c.emit_end(0);
    make_i32(&mut c, 7);
    c.emit_op(Op::RETURN, 0);
    run_chunk_expect_i32(c, 5);
}

#[test]
fn nested_if_inner_false() {
    // if (1) { if (0) { return 5 } end; return 6 } end; return 7
    let mut c = Chunk::new("<script>");
    make_i32(&mut c, 1);
    c.emit_if(0);
    make_i32(&mut c, 0);
    c.emit_if(0);
    make_i32(&mut c, 5);
    c.emit_op(Op::RETURN, 0);
    c.emit_end(0);
    make_i32(&mut c, 6);
    c.emit_op(Op::RETURN, 0);
    c.emit_end(0);
    make_i32(&mut c, 7);
    c.emit_op(Op::RETURN, 0);
    run_chunk_expect_i32(c, 6);
}

// ── VM execution: block + br ──────────────────────────────────────────────

#[test]
fn block_br_exits_block() {
    // block { br 0; [unreachable] } end; return 99
    let mut c = Chunk::new("<script>");
    c.emit_block(0);
    c.emit_br(0, 0);
    make_i32(&mut c, 0); // dead code
    c.emit_end(0);
    make_i32(&mut c, 99);
    c.emit_op(Op::RETURN, 0);
    run_chunk_expect_i32(c, 99);
}

#[test]
fn block_br_skips_unreachable_tail() {
    let mut c = Chunk::new("<script>");
    c.emit_block(0);
    c.emit_br(0, 0);
    c.emit_op(Op::UNREACHABLE, 0);
    c.emit_end(0);
    make_i32(&mut c, 99);
    c.emit_op(Op::RETURN, 0);
    run_chunk_expect_i32(c, 99);
}

#[test]
fn block_br_if_exits_when_cond_true() {
    // block { br_if(1) 0; return 10 } end; return 20
    let mut c = Chunk::new("<script>");
    c.emit_block(0);
    make_i32(&mut c, 1); // condition = true
    c.emit_br_if(0, 0);
    make_i32(&mut c, 10);
    c.emit_op(Op::RETURN, 0);
    c.emit_end(0);
    make_i32(&mut c, 20);
    c.emit_op(Op::RETURN, 0);
    run_chunk_expect_i32(c, 20);
}

#[test]
fn block_br_if_falls_through_when_cond_false() {
    // block { br_if(0) 0; return 10 } end; return 20
    let mut c = Chunk::new("<script>");
    c.emit_block(0);
    make_i32(&mut c, 0); // condition = false
    c.emit_br_if(0, 0);
    make_i32(&mut c, 10);
    c.emit_op(Op::RETURN, 0);
    c.emit_end(0);
    make_i32(&mut c, 20);
    c.emit_op(Op::RETURN, 0);
    run_chunk_expect_i32(c, 10);
}

#[test]
fn nested_blocks_br_depth_1_exits_outer() {
    // block outer { block inner { br 1 } end; return 10 } end; return 20
    let mut c = Chunk::new("<script>");
    c.emit_block(0); // outer
    c.emit_block(0); // inner
    c.emit_br(1, 0); // br 1 → exit outer
    c.emit_end(0); // end inner
    make_i32(&mut c, 10);
    c.emit_op(Op::RETURN, 0);
    c.emit_end(0); // end outer
    make_i32(&mut c, 20);
    c.emit_op(Op::RETURN, 0);
    run_chunk_expect_i32(c, 20);
}

// ── VM execution: loop + br ───────────────────────────────────────────────

#[test]
fn loop_counts_to_five() {
    // local 0 = counter (starts 0)
    // block { loop { if count >= 5, br 1; count++; br 0 } end } end
    // return count
    let mut c = Chunk::new("<script>");
    c.local_count = 1;

    c.emit_block(0); // outer block — br 1 from inside loop exits here
    c.emit_loop_s(0); // loop — br 0 restarts here

    // if count >= 5, exit
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    make_i32(&mut c, 5);
    c.emit_op(Op::I32_GE_S, 0);
    c.emit_br_if(1, 0); // br_if 1 → exit outer block

    // count++
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    make_i32(&mut c, 1);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_SET, 0, 0);

    c.emit_br(0, 0); // restart loop
    c.emit_end(0); // end loop
    c.emit_end(0); // end block

    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_op(Op::RETURN, 0);
    run_chunk_expect_i32(c, 5);
}

#[test]
fn loop_exits_early_via_block_br() {
    // count from 0, exit when count == 3
    let mut c = Chunk::new("<script>");
    c.local_count = 1;

    c.emit_block(0);
    c.emit_loop_s(0);

    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    make_i32(&mut c, 3);
    c.emit_op(Op::I32_EQ, 0);
    c.emit_br_if(1, 0); // exit block when count == 3

    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    make_i32(&mut c, 1);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_SET, 0, 0);
    c.emit_br(0, 0);
    c.emit_end(0);
    c.emit_end(0);

    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_op(Op::RETURN, 0);
    run_chunk_expect_i32(c, 3);
}

#[test]
fn nested_loops_inner_and_outer() {
    // outer_count=0, inner_count=0
    // outer_loop: 3 iterations, each with inner_loop: 2 iterations
    // total inner iterations = 6 → inner_count
    let mut c = Chunk::new("<script>");
    c.local_count = 3; // 0=outer, 1=inner, 2=total

    // outer block+loop
    c.emit_block(0);
    c.emit_loop_s(0);

    // exit outer when outer >= 3
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    make_i32(&mut c, 3);
    c.emit_op(Op::I32_GE_S, 0);
    c.emit_br_if(1, 0);

    // reset inner
    make_i32(&mut c, 0);
    c.emit_op_u16(Op::LOCAL_SET, 1, 0);

    // inner block+loop
    c.emit_block(0);
    c.emit_loop_s(0);

    // exit inner when inner >= 2
    c.emit_op_u16(Op::LOCAL_GET, 1, 0);
    make_i32(&mut c, 2);
    c.emit_op(Op::I32_GE_S, 0);
    c.emit_br_if(1, 0);

    // total++
    c.emit_op_u16(Op::LOCAL_GET, 2, 0);
    make_i32(&mut c, 1);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_SET, 2, 0);

    // inner++
    c.emit_op_u16(Op::LOCAL_GET, 1, 0);
    make_i32(&mut c, 1);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_SET, 1, 0);

    c.emit_br(0, 0); // restart inner
    c.emit_end(0); // end inner loop
    c.emit_end(0); // end inner block

    // outer++
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    make_i32(&mut c, 1);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_SET, 0, 0);

    c.emit_br(0, 0); // restart outer
    c.emit_end(0); // end outer loop
    c.emit_end(0); // end outer block

    c.emit_op_u16(Op::LOCAL_GET, 2, 0);
    c.emit_op(Op::RETURN, 0);
    run_chunk_expect_i32(c, 6);
}

// ── VM execution: br_table ────────────────────────────────────────────────

#[test]
fn br_table_selects_correct_label() {
    // block a { block b { block c {
    //   br_table [0, 1, 2] default=0 with index=2
    //   → labels[2] = 2 → depth 2 → exits a → continues to return 30
    // } end; return 10 } end; return 20 } end; return 30
    let mut c = Chunk::new("<script>");
    c.emit_block(0); // a  (depth 2 from innermost)
    c.emit_block(0); // b  (depth 1)
    c.emit_block(0); // c  (depth 0)
    make_i32(&mut c, 2); // index
    c.emit_br_table(&[0, 1, 2], 0, 0); // index=2 → depth 2 → exit a
    c.emit_end(0); // end c
    make_i32(&mut c, 10);
    c.emit_op(Op::RETURN, 0);
    c.emit_end(0); // end b
    make_i32(&mut c, 20);
    c.emit_op(Op::RETURN, 0);
    c.emit_end(0); // end a
    make_i32(&mut c, 30);
    c.emit_op(Op::RETURN, 0);
    run_chunk_expect_i32(c, 30);
}

#[test]
fn br_table_uses_default_for_out_of_range() {
    // br_table [0] default=1 with index=5 (out of range → default=1)
    let mut c = Chunk::new("<script>");
    c.emit_block(0); // outer (depth 1)
    c.emit_block(0); // inner (depth 0)
    make_i32(&mut c, 5); // out of range
    c.emit_br_table(&[0], 1, 0); // default depth=1 → exit outer
    c.emit_end(0); // end inner
    make_i32(&mut c, 10);
    c.emit_op(Op::RETURN, 0);
    c.emit_end(0); // end outer
    make_i32(&mut c, 20);
    c.emit_op(Op::RETURN, 0);
    run_chunk_expect_i32(c, 20);
}

// ── VM execution: if-value (result-bearing if/else) ───────────────────────

#[test]
fn if_value_true_leaves_correct_result() {
    let mut c = Chunk::new("<script>");
    make_i32(&mut c, 1);
    c.emit_if_value(0);
    make_i32(&mut c, 100);
    c.emit_else(0);
    make_i32(&mut c, 200);
    c.emit_end(0);
    c.emit_op(Op::RETURN, 0);
    run_chunk_expect_i32(c, 100);
}

#[test]
fn if_value_false_leaves_correct_result() {
    let mut c = Chunk::new("<script>");
    make_i32(&mut c, 0);
    c.emit_if_value(0);
    make_i32(&mut c, 100);
    c.emit_else(0);
    make_i32(&mut c, 200);
    c.emit_end(0);
    c.emit_op(Op::RETURN, 0);
    run_chunk_expect_i32(c, 200);
}

// ── VM execution: loop as do-while ───────────────────────────────────────

#[test]
fn do_while_loop_executes_body_at_least_once() {
    // Execute body once even though condition is false from the start.
    // count = 0; do { count++; } while count < 0
    // Result: count = 1
    let mut c = Chunk::new("<script>");
    c.local_count = 1;

    c.emit_block(0);
    c.emit_loop_s(0);

    // body: count++
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    make_i32(&mut c, 1);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_SET, 0, 0);

    // continue while count < 0 (immediately false → exit)
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    make_i32(&mut c, 0);
    c.emit_op(Op::I32_LT_S, 0);
    c.emit_br_if(0, 0); // br_if 0 → restart loop (not taken)

    c.emit_end(0); // end loop
    c.emit_end(0); // end block

    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_op(Op::RETURN, 0);
    run_chunk_expect_i32(c, 1);
}

// ── VM execution: continue-equivalent (br 0 skips rest of loop body) ─────

#[test]
fn loop_continue_skips_sum_for_odd_numbers() {
    // Sum even numbers 0..5: 0+2+4 = 6
    // i=0; sum=0
    // loop { if i>=5, exit; if i%2!=0, i++, continue; sum+=i; i++; restart }
    let mut c = Chunk::new("<script>");
    c.local_count = 2; // 0=i, 1=sum

    c.emit_block(0); // exit block
    c.emit_loop_s(0); // loop

    // exit when i >= 5
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    make_i32(&mut c, 5);
    c.emit_op(Op::I32_GE_S, 0);
    c.emit_br_if(1, 0);

    // if i is odd (i & 1 != 0): i++, br 0 (continue)
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    make_i32(&mut c, 1);
    c.emit_op(Op::I32_AND, 0); // i & 1
    c.emit_if(0);
    // i++
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    make_i32(&mut c, 1);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_SET, 0, 0);
    c.emit_br(1, 0); // br 1 restarts loop (1=loop inside outer if, relative depth)
    c.emit_end(0); // end if

    // sum += i
    c.emit_op_u16(Op::LOCAL_GET, 1, 0);
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_SET, 1, 0);

    // i++
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    make_i32(&mut c, 1);
    c.emit_op(Op::I32_ADD, 0);
    c.emit_op_u16(Op::LOCAL_SET, 0, 0);

    c.emit_br(0, 0); // restart loop
    c.emit_end(0); // end loop
    c.emit_end(0); // end block

    c.emit_op_u16(Op::LOCAL_GET, 1, 0);
    c.emit_op(Op::RETURN, 0);
    run_chunk_expect_i32(c, 6);
}

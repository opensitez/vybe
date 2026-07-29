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
            .windows(4)
            .any(|w| w == Op::IF.encode())
    );
    assert!(
        chunk
            .code
            .windows(4)
            .any(|w| w == Op::ELSE.encode())
    );
    assert!(
        chunk
            .code
            .windows(4)
            .any(|w| w == Op::END.encode())
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
            .windows(4)
            .any(|w| w == Op::BR_IF.encode())
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
        .windows(4)
        .position(|w| w == Op::BR_TABLE.encode())
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
            .windows(4)
            .any(|w| w == [Op::I31_NEW.group(), Op::I31_NEW.sub()])
    );
    assert!(
        chunk
            .code
            .windows(4)
            .any(|w| w == [Op::V128_CONST.group(), Op::V128_CONST.sub()])
    );
    assert!(
        chunk
            .code
            .windows(4)
            .any(|w| w == Op::ATOMIC_FENCE.encode())
    );
}


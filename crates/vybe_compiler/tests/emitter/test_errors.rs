use vybe_compiler::primitives::errors;
use vybe_runtime::Chunk;

#[test]
fn emit_try_start_opens_handler_block_then_try_table() {
    use vybe_runtime::opcode::Op;
    let mut chunk = Chunk::new("test");
    errors::emit_try_start(&mut chunk, 0);

    // A handler BLOCK comes first: the clause's `labelidx` needs a block to
    // name, and its `end` is where the catch arms begin.
    let first = Op::decode(
        ((chunk.code[0] as u16) << 8) | chunk.code[1] as u16,
        ((chunk.code[2] as u16) << 8) | chunk.code[3] as u16,
    );
    assert_eq!(first, Some(Op::BLOCK), "try_start must open a handler block");
    // It carries ONE result — the exception object travels as the block's
    // result when the clause branches.
    assert_eq!(chunk.code[4], 0, "handler block takes no params");
    assert_eq!(chunk.code[5], 1, "handler block must carry one result");

    let second = Op::decode(
        ((chunk.code[6] as u16) << 8) | chunk.code[7] as u16,
        ((chunk.code[8] as u16) << 8) | chunk.code[9] as u16,
    );
    assert_eq!(second, Some(Op::TRY_TABLE));
    assert_eq!(chunk.code[10], 1, "one catch clause");
    // [kind, tag hi, tag lo, label hi, label lo] — the label is 0, naming the
    // handler block just opened. This is a labelidx (a block depth), NOT a byte
    // offset; nothing is patched into it.
    assert_eq!(
        (chunk.code[14], chunk.code[15]),
        (0, 0),
        "clause must carry labelidx 0 — the handler block"
    );
}

#[test]
fn emit_try_end_produces_bytecode() {
    let mut chunk = Chunk::new("test");
    errors::emit_try_end(&mut chunk, 0);
    assert!(!chunk.code.is_empty(), "try_end should emit opcode");
}

#[test]
fn emit_throw_produces_bytecode() {
    let mut chunk = Chunk::new("test");
    errors::emit_throw(&mut chunk, 0);
    assert!(!chunk.code.is_empty(), "throw should emit opcode");
}

#[test]
fn try_catch_roundtrip_is_balanced_and_patch_free() {
    use vybe_runtime::opcode::Op;
    let mut chunk = Chunk::new("test");
    errors::emit_try_start(&mut chunk, 0);
    // ... body would go here ...
    errors::emit_try_end(&mut chunk, 0);
    chunk.emit_br(1, 0); // normal path: past the handler
    errors::emit_handler_block_end(&mut chunk, 0);
    // ... handler would go here ...

    // Two openers (handler block + try_table) and two `end`s. An unbalanced
    // region is what a missing `end` would leave, and it would make
    // `build_block_table` pair the try_table with the WRONG `end`.
    let mut openers = 0;
    let mut ends = 0;
    let mut ip = 0;
    while ip + 3 < chunk.code.len() {
        let op = Op::decode(
            ((chunk.code[ip] as u16) << 8) | chunk.code[ip + 1] as u16,
            ((chunk.code[ip + 2] as u16) << 8) | chunk.code[ip + 3] as u16,
        );
        match op {
            Some(Op::BLOCK) | Some(Op::TRY_TABLE) => openers += 1,
            Some(Op::END) => ends += 1,
            _ => {}
        }
        ip += 4 + op.map_or(0, |o| o.operand_format().size_in(&chunk.code, ip + 4));
    }
    assert_eq!(openers, 2, "handler block + try_table");
    assert_eq!(ends, 2, "each must be closed");
}

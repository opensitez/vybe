use vybe_runtime::link_preparation::PreparedLinkedCode;
use vybe_runtime::{Chunk, Op, VM, Value};

#[test]
fn compact_opcode_metadata_respects_budget_and_preserves_owned_code() {
    let mut chunk = Chunk::new("<opcode-bucket-budget>");
    for value in 0..2048 {
        chunk.emit_i32_const(value, 17);
        chunk.emit_op(Op::DROP, 17);
    }
    chunk.emit_i32_const(7, 17);
    chunk.emit_op(Op::RETURN, 17);
    let bytes = chunk.code.clone();
    // The old byte-indexed opcode vector alone consumed four bytes per
    // bytecode byte, plus three more for operand caches. Compact metadata
    // must fit below that footprint without changing the emitted program.
    let budget = bytes.len() * 5;
    let Ok(prepared) = PreparedLinkedCode::prepare(vec![chunk], budget) else {
        panic!("opcode metadata exceeded its compact storage budget");
    };
    assert!(prepared.metadata_bytes() <= budget);
    assert_eq!(prepared.chunks()[0].code, bytes);
    let raw = prepared.into_raw();
    assert_eq!(raw[0].code, bytes);
    let mut vm = VM::new();
    assert_eq!(vm.run_linked_nested(raw, vec![]).unwrap(), Value::I32(7));
}

use super::*;

fn local(chunk: &mut Chunk, op: Op, slot: u16) {
    chunk.emit_op(op, 0);
    chunk.code.extend_from_slice(&slot.to_be_bytes());
    chunk.lines.resize(chunk.code.len(), 0);
}

#[test]
fn self_copy_keeps_both_type_observations_when_recording() {
    let mut chunk = Chunk::new("self-copy-recording");
    chunk.local_count = 1;
    chunk.scratch_high_water = 1;
    chunk.result_arity = 1;
    chunk.emit_i32_const(42, 0);
    local(&mut chunk, Op::LOCAL_SET, 0);
    for _ in 0..3 {
        local(&mut chunk, Op::LOCAL_GET, 0);
        local(&mut chunk, Op::LOCAL_SET, 0);
    }
    local(&mut chunk, Op::LOCAL_GET, 0);
    chunk.emit_op(Op::RETURN, 0);
    let mut vm = VM::new();
    vm.type_recorder = Some(crate::type_recorder::TypeRecorder::new());
    assert_eq!(vm.run(vec![chunk]).unwrap(), Value::I32(42));
    let recorder = vm.type_recorder.as_ref().unwrap();
    assert_eq!(recorder.slots()[0][0].total(), 8);
    assert_eq!(recorder.slots()[0][0].distinct_variants(), 1);
}

#[test]
fn invalid_self_copy_does_not_bypass_local_validation() {
    let mut chunk = Chunk::new("invalid-self-copy");
    chunk.local_count = 1;
    chunk.scratch_high_water = 1;
    chunk.result_arity = 1;
    local(&mut chunk, Op::LOCAL_GET, 1);
    local(&mut chunk, Op::LOCAL_SET, 1);
    chunk.emit_i32_const(42, 0);
    chunk.emit_op(Op::RETURN, 0);
    assert!(VM::new().run(vec![chunk]).is_err());
}

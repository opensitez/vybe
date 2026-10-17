use vybe_compiler::primitives::platforms::register_platforms;
use vybe_runtime::capabilities::Capabilities;
use vybe_runtime::value::{ObjectKind, Value};
use vybe_runtime::{Chunk, Op, VM};

#[test]
fn request_constructor_retains_body_written_after_construction() {
    let mut chunk = Chunk::new("<http-deferred-body>");
    let fields = chunk.add_import("wasi:http/types", "[constructor]fields");
    let request_new = chunk.add_import("wasi:http/types", "[static]request.new");
    let consume = chunk.add_import("wasi:http/types", "[static]request.consume-body");
    let at = chunk.add_import("ecma:array", "at");
    let new = chunk.add_import("canon", "stream.new");
    let write = chunk.add_import("canon", "stream.write");
    let drop_write = chunk.add_import("canon", "stream.drop-writable");
    let reader = chunk.alloc_scratch(1);
    let writer = chunk.alloc_scratch(1);
    let request = chunk.alloc_scratch(1);
    chunk.local_count = chunk.local_count.max(chunk.scratch_high_water);
    let packed = chunk.i64_scratch();

    chunk.emit_call(new, 0, 0);
    chunk.emit_op_u16(Op::LOCAL_TEE, packed, 0);
    chunk.emit_i64_const(32, 0);
    chunk.emit_op(Op::I64_SHR_U, 0);
    chunk.emit_op(Op::I32_WRAP_I64, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, writer, 0);
    chunk.emit_op_u16(Op::LOCAL_GET, packed, 0);
    chunk.emit_op(Op::I32_WRAP_I64, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, reader, 0);

    chunk.emit_call(fields, 0, 0);
    chunk.emit_op_u16(Op::LOCAL_GET, reader, 0);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    chunk.emit_call(request_new, 4, 0);
    chunk.emit_op_u16(Op::LOCAL_SET, request, 0);

    let (ptr, len) = vybe_compiler::primitives::canon_marshal::emit_store_utf8(
        &mut chunk, 0, |chunk| { chunk.emit_string_const("later", 0); },
    );
    chunk.emit_op_u16(Op::LOCAL_GET, writer, 0);
    chunk.emit_op_u16(Op::LOCAL_GET, ptr, 0);
    chunk.emit_op_u16(Op::LOCAL_GET, len, 0);
    chunk.emit_call(write, 3, 0);
    chunk.emit_op(Op::DROP, 0);
    chunk.emit_op_u16(Op::LOCAL_GET, writer, 0);
    chunk.emit_call(drop_write, 1, 0);

    chunk.emit_op_u16(Op::LOCAL_GET, request, 0);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
    chunk.emit_call(consume, 2, 0);
    chunk.emit_i32_const(0, 0);
    chunk.emit_call(at, 2, 0);
    vybe_compiler::primitives::io::emit_read_stream_to_bytes(&mut chunk, 0);
    chunk.emit_op(Op::RETURN, 0);

    let mut vm = VM::new();
    register_platforms(&mut vm, &Capabilities::all());
    let result = vm.run(vec![chunk]).expect("deferred HTTP body caller");
    let Value::Object(bytes) = result else { panic!("expected body bytes") };
    let bytes = bytes.lock().unwrap();
    let ObjectKind::Array(bytes) = &bytes.kind else { panic!("expected byte array") };
    assert_eq!(bytes.iter().map(|byte| byte.as_i32() as u8).collect::<Vec<_>>(), b"later");
}

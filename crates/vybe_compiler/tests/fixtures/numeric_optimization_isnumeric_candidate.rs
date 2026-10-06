// Isolated primitive-tag predicates; production unchanged.
use vybe_runtime::{Chunk, Op};
pub fn build_is_numeric(_imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_isnumeric");
    c.arity = 1;
    c.local_count = 2;
    let test_number = c.add_import("wasm:js-number", "test");
    let test_string = c.add_import("wasm:js-string", "test");
    let parse = c.add_import("ecma:number", "parseFloat");
    let done = c.emit_block(0);
    // Same numeric tag set as ecma:value.typeof == "number", including
    // F32/I32/I64 and NaN. Result remains the original boxed-I320/1 ABI.
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_call(test_number, 1, 0);
    c.emit_op_u16(Op::LOCAL_SET, 1, 0);
    c.emit_op_u16(Op::LOCAL_GET, 1, 0);
    c.emit_br_if(0, 0);
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_call(test_string, 1, 0);
    c.emit_op(Op::I32_EQZ, 0);
    c.emit_br_if(0, 0); // all other primitive/object tags retain false
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_call(parse, 1, 0);
    c.emit_dup(0);
    c.emit_op(Op::F64_EQ, 0);
    c.emit_op_u16(Op::LOCAL_SET, 1, 0);
    c.emit_end(0);
    c.patch_block(done);
    c.emit_op_u16(Op::LOCAL_GET, 1, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

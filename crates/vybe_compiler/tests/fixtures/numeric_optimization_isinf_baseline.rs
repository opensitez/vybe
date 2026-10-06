// Exact pre-specialization build_isinf factory, 2026-10-04.
use vybe_runtime::{Chunk, Op};
pub fn build_isinf(imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_isinf");
    c.arity = 1;
    c.local_count = 1;
    let isfin = c.add_import("ecma:number", "isFinite");
    let isnan = c.add_import("ecma:number", "isNaN");

    // !isFinite(n) && !isNaN(n)
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_call(isfin, 1, 0);
    vybe_compiler::primitives::ops::emit_dyn_not_into(imports, &mut c, 0);
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_call(isnan, 1, 0);
    vybe_compiler::primitives::ops::emit_dyn_not_into(imports, &mut c, 0);
    c.emit_op(Op::I32_AND, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

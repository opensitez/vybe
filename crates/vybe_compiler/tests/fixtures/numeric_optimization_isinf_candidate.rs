// Isolated rejected checked-Boolean candidate, 2026-10-04.
// Infinity regressed in cold and loaded opt-level-0 measurements; production
// retains its original factory. This fixture preserves the experiment only.
use vybe_runtime::{Chunk, Op};
pub fn build_isinf(_imports: &mut Chunk) -> Chunk {
    let mut c = Chunk::new("__stdlib_isinf");
    c.arity = 1;
    c.local_count = 1;
    let isfin = c.add_import("ecma:number", "isFinite");
    let isnan = c.add_import("ecma:number", "isNaN");
    // Both strict Number predicates declare and produce Boolean on every arm.
    // Convert those checked Booleans directly; preserve their strict input rules.
    let cast_bool = c.add_import("wasm:js-boolean", "cast");

    // !isFinite(n) && !isNaN(n)
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_call(isfin, 1, 0);
    c.emit_call(cast_bool, 1, 0);
    c.emit_op(Op::I32_EQZ, 0);
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    c.emit_call(isnan, 1, 0);
    c.emit_call(cast_bool, 1, 0);
    c.emit_op(Op::I32_EQZ, 0);
    c.emit_op(Op::I32_AND, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

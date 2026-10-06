// Isolated candidate: unchanged actual toString producer; proven string concat.
use vybe_runtime::{Chunk, Op, Value};
pub fn build_pyradix(_imports: &mut Chunk, name: &str, prefix: &str, radix: i32) -> Chunk {
    let mut c = Chunk::new(name);
    c.arity = 1;
    c.local_count = 1;
    let pref = c.add_constant(Value::String(std::sync::Arc::from(prefix)));
    let r = c.add_constant(Value::I32(radix));
    let ts_idx = c.add_import("ecma:number", "toString");
    let Value::String(prefix) = c.constants[pref as usize].clone() else {
        unreachable!()
    };
    c.emit_string_const(&prefix, 0);
    c.emit_op_u16(Op::LOCAL_GET, 0, 0);
    let Value::I32(radix) = c.constants[r as usize] else {
        unreachable!()
    };
    c.emit_i32_const(radix, 0);
    c.emit_call(ts_idx, 2, 0);
    let concat = c.add_import("wasm:js-string", "concat");
    c.emit_call(concat, 2, 0);
    c.emit_op(Op::RETURN, 0);
    c
}

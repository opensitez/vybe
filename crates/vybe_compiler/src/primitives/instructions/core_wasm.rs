use vybe_runtime::Value;
use vybe_runtime::{Chunk, Op};

pub fn dup(c: &mut Chunk, line: u32) {
    c.emit_dup(line);
}

pub fn i32_const(c: &mut Chunk, line: u32, v: i32) {
    c.emit_i32_const(v, line);
}

pub fn i64_const(c: &mut Chunk, line: u32, v: i64) {
    c.emit_i64_const(v, line);
}

pub fn f64_const(c: &mut Chunk, line: u32, v: f64) {
    c.emit_f64_const(v, line);
}

pub fn bool_const(c: &mut Chunk, line: u32, v: bool) {
    c.emit_bool_const(v, line);
}

pub fn undefined(c: &mut Chunk, line: u32) {
    crate::primitives::globals::emit_read(c, "undefined", line);
}

pub fn null(c: &mut Chunk, line: u32) {
    c.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

pub fn string_const(c: &mut Chunk, line: u32, s: &str) {
    c.emit_string_const(s, line);
}

pub fn emit_value(c: &mut Chunk, line: u32, val: &Value) {
    match val {
        Value::I32(v) => c.emit_i32_const(*v, line),
        Value::I64(v) => c.emit_i64_const(*v, line),
        Value::F64(v) => c.emit_f64_const(*v, line),
        Value::Bool(v) => c.emit_bool_const(*v, line),
        Value::Null => c.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line),
        Value::Undefined => undefined(c, line),
        Value::String(s) => c.emit_string_const(s, line),
        // Language BigInts belong to the boxed value ABI. A raw i64.const
        // cannot feed externref consumers, and wrapping loses larger values.
        // Explicit Wasm i64 constants still use i64_const above.
        Value::BigInt(v) => {
            c.emit_string_const(&v.to_string(), line);
            let ctor = c.add_import("ecma:bigint", "BigInt");
            c.emit_call(ctor, 1, line);
        }
        Value::V128(v) => {
            c.emit_op(Op::V128_CONST, line);
            for b in v {
                c.emit(*b, line);
            }
        }
        other => panic!("emit_value: no WASM-compliant encoding for {:?}", other),
    }
}

/// Store the top value while retaining it as the next operand.
/// Used only where the original local.set/local.get were adjacent emissions.
pub fn local_tee(c: &mut Chunk, line: u32, slot: u16) {
    #[cfg(feature = "primitive-benchmarks")]
    if !REUSE_LOCAL_TEE.with(|v| v.get()) {
        c.emit_op_u16(Op::LOCAL_SET, slot, line);
        c.emit_op_u16(Op::LOCAL_GET, slot, line);
        return;
    }
    c.emit_op_u16(Op::LOCAL_TEE, slot, line);
}

#[cfg(feature = "primitive-benchmarks")]
thread_local! {
    static REUSE_LOCAL_TEE: std::cell::Cell<bool> = const { std::cell::Cell::new(true) };
}

#[cfg(feature = "primitive-benchmarks")]
pub fn benchmark_with_local_tee<R>(reuse: bool, build: impl FnOnce() -> R) -> R {
    struct Restore(bool);
    impl Drop for Restore {
        fn drop(&mut self) { REUSE_LOCAL_TEE.with(|v| v.set(self.0)); }
    }
    let _restore = Restore(REUSE_LOCAL_TEE.with(|v| v.replace(reuse)));
    build()
}

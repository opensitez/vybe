use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;
use vybe_runtime::opcode::heaptype::{HT_FUNC, HT_STRUCT, HeapType};

use super::host::{EmitRegistry, FunctionRegistry};

pub fn register_all(_fns: &FunctionRegistry, emits: &mut EmitRegistry) {
    emits.register("common:is_object", is_object_fn);
    emits.register("common:is_func", is_func_fn);
}

fn is_object_fn(_fns: &FunctionRegistry, c: &mut Chunk, line: u32) {
    is_object(c, line);
}

fn is_func_fn(_fns: &FunctionRegistry, c: &mut Chunk, line: u32) {
    is_func(c, line);
}

/// `ref.test struct` — "is this an object?" is a question about the ABSTRACT
/// heap hierarchy, not about a type called `Object`. Asking it as a concrete
/// test is what forced every user class to declare `Object` as its supertype,
/// because otherwise the registry walk found nothing.
pub fn is_object(c: &mut Chunk, line: u32) {
    c.emit_ref_type_op(Op::REF_TEST, HeapType::Abstract(HT_STRUCT), line);
}

/// `ref.test func` — top of the function hierarchy.
pub fn is_func(c: &mut Chunk, line: u32) {
    c.emit_ref_type_op(Op::REF_TEST, HeapType::Abstract(HT_FUNC), line);
}

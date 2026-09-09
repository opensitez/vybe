//! PHP JSON adapter helpers.
//!
//! Public JSON builtins live in their existing adapters; this file owns the
//! PHP-specific normalization step used before JSON encoding.

use vybe_runtime::opcode::Op;
use vybe_runtime::Chunk;

fn lset(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}

/// Build and call the common JSON normalizer in PHP mode.
pub fn emit_php_json_normalize(
    chunks: &mut Vec<Chunk>,
    current: usize,
    value_slot: u16,
    line: u32,
) {
    let (default_slot, sort_slot, props_slot) = {
        let c = &mut chunks[current];
        (c.alloc_scratch(1), c.alloc_scratch(1), c.alloc_scratch(1))
    };
    {
        let c = &mut chunks[current];
        c.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        lset(c, default_slot, line);
        c.emit_bool_const(false, line);
        lset(c, sort_slot, line);
        c.emit_bool_const(true, line);
        lset(c, props_slot, line);
    }
    vybe_compiler::primitives::json::emit_normalize(
        chunks,
        current,
        value_slot,
        default_slot,
        sort_slot,
        props_slot,
        line,
    );
}

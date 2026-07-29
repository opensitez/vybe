//! Python `json.dumps` — Python semantics over the shared `vybe_compiler::primitives::json`
//! core.
//!
//! The walker reshapes `json.dumps(obj, cls=…, default=…, sort_keys=…,
//! indent=…, separators=…)` into the fixed positional form
//! `__py_json_dumps(value, default, sort_keys, indent, item_sep, kv_sep)`
//! (see `walker::rewrite_json_dumps`). This adapter normalizes the value tree
//! (applying the `default`/`cls` encoder hook to non-serializable objects like
//! `datetime`), then renders it: indented output goes through
//! `ecma:json.stringify` (byte-identical to Python's), compact output through
//! the shared separator-aware renderer so Python's `", "` / `": "` defaults
//! come out right.

use vybe_bytecode::Chunk;
use vybe_bytecode::opcode::Op;

/// `emit = "common:python.json_dumps"`.
/// Stack in (bottom→top): value, default, sort_keys, indent, item_sep, kv_sep.
/// Leaves the JSON string on the stack.
pub fn emit_json_dumps(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    // Self-sufficient target: the walker's rewrite supplies all six, but a
    // BINDING (`from json import dumps`) calls it with only the positional
    // args the user wrote. Fill the missing trailing slots with Python's own
    // defaults so the adapter behaves the same either way — that is what lets
    // a name bound to this target match `json.dumps(...)` exactly.
    if argc < 6 {
        let c = &mut chunks[current];
        // default=None, sort_keys=False, indent=None, separators=(", ", ": ")
        for i in argc..6 {
            match i {
                1 | 3 => {
                    let k = c.add_constant(vybe_bytecode::Value::Null);
                    c.emit_op_u16(Op::CONST, k, line);
                }
                2 => c.emit_i32_const(0, line),
                4 => c.emit_string_const(", ", line),
                _ => c.emit_string_const(": ", line),
            }
        }
    }

    let (value_slot, default_slot, sort_slot, indent_slot, item_slot, kv_slot, props_slot, norm_slot) = {
        let c = &mut chunks[current];
        (
            c.alloc_scratch(1),
            c.alloc_scratch(1),
            c.alloc_scratch(1),
            c.alloc_scratch(1),
            c.alloc_scratch(1),
            c.alloc_scratch(1),
            c.alloc_scratch(1),
            c.alloc_scratch(1),
        )
    };

    {
        let c = &mut chunks[current];
        // Pop the six args (top → bottom).
        c.emit_op_u16(Op::LOCAL_SET, kv_slot, line);
        c.emit_op_u16(Op::LOCAL_SET, item_slot, line);
        c.emit_op_u16(Op::LOCAL_SET, indent_slot, line);
        c.emit_op_u16(Op::LOCAL_SET, sort_slot, line);
        c.emit_op_u16(Op::LOCAL_SET, default_slot, line);
        c.emit_op_u16(Op::LOCAL_SET, value_slot, line);
        // props = false — Python routes class instances to the encoder hook.
        c.emit_bool_const(false, line);
        c.emit_op_u16(Op::LOCAL_SET, props_slot, line);
    }

    // normalized = normalize(value, default, sort_keys, props=false)
    vybe_compiler::primitives::json::emit_normalize(
        chunks,
        current,
        value_slot,
        default_slot,
        sort_slot,
        props_slot,
        line,
    );
    chunks[current].emit_op_u16(Op::LOCAL_SET, norm_slot, line);

    // indent is not None → ecma:json.stringify(normalized, null, indent), whose
    // indented output is byte-identical to Python's. Otherwise render with the
    // (Python-default or caller-supplied) compact separators.
    {
        let c = &mut chunks[current];
        c.emit_op_u16(Op::LOCAL_GET, indent_slot, line);
        c.emit_op(Op::REF_IS_NULL, line);
        c.emit_op(Op::I32_EQZ, line);
        c.emit_if_value(line);
        c.emit_op_u16(Op::LOCAL_GET, norm_slot, line);
        c.emit_op(Op::NULL, line);
        c.emit_op_u16(Op::LOCAL_GET, indent_slot, line);
        let idx = c.add_import("ecma:json", "stringify");
        c.emit_call(idx, 3, line);
        c.emit_else(line);
    }
    vybe_compiler::primitives::json::emit_render_separated(chunks, current, norm_slot, item_slot, kv_slot, line);
    chunks[current].emit_end(line);
}

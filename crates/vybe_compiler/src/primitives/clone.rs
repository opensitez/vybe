//! Copying a value — the one place that answers "what does `b = a` duplicate".
//!
//! Six implementations of this question existed before this module, and none
//! of them could see each other:
//!
//! | where | depth | rtt | reach |
//! |---|---|---|---|
//! | `records::emit_value_copy` (assignment) | deep | lost | runtime stamp |
//! | `type_inference::emit_user_value_type_clone_from_stack` (arguments) | shallow | kept | compile-time only |
//! | `collections::emit_clone` | shallow | — | list/map `.clone()` |
//! | `type_inference::emit_array_clone_from_stack` | shallow | — | arrays |
//! | `structuredClone` host fn | deep | — | pascal, JS, python `deepcopy` |
//! | `__php_copy_on_assign` prelude | deep, arrays only | — | PHP |
//!
//! Two of them disagreed about whether the copy keeps its type identity, and
//! only the two deep ones could handle a value type holding a value type.
//!
//! **`ProtocolSlot::Clone` is the interop channel and it already exists.** Nine
//! languages map their spelling onto it — C#/Go `Clone`, java `clone`, php
//! `__clone`, python `__copy__`, ruby `initialize_copy`, vb/powershell `clone`,
//! pascal `assign` — and `protocol_slot_key` writes it on the INSTANCE under a
//! language-neutral name, the same way `BinOp::Eq` reaches python's `__eq__`
//! from any language. Nothing read it. Reading it here is what makes PHP able
//! to clone a Python object.
//!
//! **Not everything deepens.** Java's `ArrayList.clone()` and python's
//! `list.copy()` are shallow BY SPEC. Depth is a parameter a caller passes, not
//! a property of this module — unifying the mechanism must not unify the
//! semantics.

use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

use crate::primitives::instructions::recipes;
use crate::primitives::{loops, ops, records};

/// The shared recursive copy helper. One chunk per compilation, keyed by name.
const COPY_CHUNK: &str = "__vybe_value_copy";

/// Copy the value in `slot`. Stack: `[] → [value]`.
///
/// `force` skips the "is this even a value?" tests because the caller already
/// answered from the STATIC type — Pascal's `var a, b: TR` default-initialises
/// and never runs a constructor, so it carries no stamp to read. It is a
/// compile-time constant at every call site, which keeps the helper
/// monomorphic; the recursive calls inside always pass `false` so a nested
/// REFERENCE stays shared.
pub fn emit_deep_copy(chunks: &mut Vec<Chunk>, current: usize, slot: u16, force: bool, line: u32) {
    let helper = ensure_copy_chunk(chunks);
    chunks[current].emit_op_u16(Op::REF_FUNC, helper as u16, line);
    chunks[current].emit(0u8, line); // upvalue count
    chunks[current].emit_op_u16(Op::LOCAL_GET, slot, line);
    chunks[current].emit_bool_const(force, line);
    chunks[current].emit_op(Op::CALL_REF, line);
    chunks[current].emit(2u8, line);
}

/// Idempotent by chunk name, the same shape `sprintf::ensure_chunk` uses.
///
/// The chunk is PUSHED before its body is emitted so the body can name its own
/// index for the recursive call — self-reference needs the index already fixed.
/// Because the body is emitted with `current = idx`, it is written with the
/// ordinary emitters rather than hand-rolled opcodes.
fn ensure_copy_chunk(chunks: &mut Vec<Chunk>) -> usize {
    if let Some(idx) = chunks.iter().position(|c| c.name == COPY_CHUNK) {
        return idx;
    }
    let mut c = Chunk::new(COPY_CHUNK);
    c.arity = 2; // (value, force)
    c.local_count = 2;
    chunks.push(c);
    let idx = chunks.len() - 1;
    emit_copy_body(chunks, idx);
    idx
}

/// `__vybe_value_copy(value, force) -> value`
///
/// ```text
/// if !force {
///     value is not an aggregate  -> value    // number, string, bool, null
///     value is a function        -> value    // closures are references
///     value has no __value_copy  -> value    // a reference object
/// }
/// out = {}; Object.assign(out, value)
/// for k in Object.keys(out): out[k] = __vybe_value_copy(out[k], false)
/// return out
/// ```
///
/// **No cycle guard**, inherited from the PHP prelude which had none either: a
/// field pointing back at its container recurses until the stack gives out.
fn emit_copy_body(chunks: &mut [Chunk], idx: usize) {
    const LINE: u32 = 0;
    let value = 0u16;
    let force = 1u16;

    let base = chunks[idx].local_count;
    chunks[idx].alloc_scratch(5);
    let (out, keys, i_slot, key, copied) =
        (base, base + 1, base + 2, base + 3, base + 4);

    chunks[idx].emit_op_u16(Op::LOCAL_GET, force, LINE);
    ops::emit_dyn_to_bool(&mut chunks[idx], LINE);
    chunks[idx].emit_op(Op::I32_EQZ, LINE);
    chunks[idx].emit_if(LINE);
    {
        let bail_if = |chunks: &mut [Chunk], test: &dyn Fn(&mut Chunk)| {
            test(&mut chunks[idx]);
            chunks[idx].emit_if(LINE);
            chunks[idx].emit_op_u16(Op::LOCAL_GET, value, LINE);
            chunks[idx].emit_op(Op::RETURN, LINE);
            chunks[idx].emit_end(LINE);
        };
        // Not a heap aggregate at all — scalars and strings are already values.
        bail_if(chunks, &|c| {
            c.emit_op_u16(Op::LOCAL_GET, value, LINE);
            recipes::is_object(c, LINE);
            c.emit_op(Op::I32_EQZ, LINE);
        });
        // A function or closure. Map-backed like everything else, but a
        // reference by nature — deep-copying one shreds it.
        bail_if(chunks, &|c| {
            c.emit_op_u16(Op::LOCAL_GET, value, LINE);
            recipes::is_func(c, LINE);
        });
        // An aggregate whose declaration never asked for value storage.
        bail_if(chunks, &|c| {
            records::emit_is_value_copy(c, value, LINE);
            c.emit_op(Op::I32_EQZ, LINE);
        });
    }
    chunks[idx].emit_end(LINE);

    // `Object.assign({}, src)` — a fresh object carrying src's own fields. The
    // same shape `emit_inherit_statics` uses, and what C#'s `with` expression
    // hand-rolled.
    chunks[idx].emit_struct_new(0, 0, LINE);
    chunks[idx].emit_op_u16(Op::LOCAL_GET, value, LINE);
    let assign_fn = chunks[idx].add_import("ecma:object", "assign");
    chunks[idx].emit_call(assign_fn, 2, LINE);
    chunks[idx].emit_op_u16(Op::LOCAL_SET, out, LINE);

    // Replace each copied field with a copy of its own, so a nested value type
    // is independent too. Iterating the COPY's keys, not the source's, keeps
    // this to whatever `assign` actually brought over.
    chunks[idx].emit_op_u16(Op::LOCAL_GET, out, LINE);
    let keys_fn = chunks[idx].add_import("ecma:object", "keys");
    chunks[idx].emit_call(keys_fn, 1, LINE);
    chunks[idx].emit_op_u16(Op::LOCAL_SET, keys, LINE);

    let state = loops::emit_for_in_start(chunks, idx, keys, i_slot, LINE);
    chunks[idx].emit_op_u16(Op::LOCAL_SET, key, LINE);
    // `out[key] = self(out[key], false)`, in two steps.
    //
    // The call is issued on a CLEAN stack and its result sunk into a local
    // before the `set` operands are pushed. Building `[out, key]` first and
    // calling with them underneath the callee is what `emit_sprintf_from_array`
    // deliberately avoids — it stashes into locals for the same reason. Doing
    // it the other way round, the recursion demonstrably re-entered (a forced
    // run overflowed the stack inside this chunk) yet its result never reached
    // the field, while writing a constant in the same position worked.
    chunks[idx].emit_op_u16(Op::REF_FUNC, idx as u16, LINE);
    chunks[idx].emit(0u8, LINE); // upvalue count
    chunks[idx].emit_op_u16(Op::LOCAL_GET, out, LINE);
    chunks[idx].emit_op_u16(Op::LOCAL_GET, key, LINE);
    let get_fn = chunks[idx].add_import("ecma:object", "get");
    chunks[idx].emit_call(get_fn, 2, LINE);
    chunks[idx].emit_bool_const(false, LINE);
    chunks[idx].emit_op(Op::CALL_REF, LINE);
    chunks[idx].emit(2u8, LINE);
    chunks[idx].emit_op_u16(Op::LOCAL_SET, copied, LINE);

    chunks[idx].emit_op_u16(Op::LOCAL_GET, out, LINE);
    chunks[idx].emit_op_u16(Op::LOCAL_GET, key, LINE);
    chunks[idx].emit_op_u16(Op::LOCAL_GET, copied, LINE);
    let set_fn = chunks[idx].add_import("ecma:object", "set");
    chunks[idx].emit_call(set_fn, 3, LINE);
    chunks[idx].emit_op(Op::DROP, LINE);
    loops::emit_for_in_end(chunks, idx, i_slot, state, LINE);

    chunks[idx].emit_op_u16(Op::LOCAL_GET, out, LINE);
    chunks[idx].emit_op(Op::RETURN, LINE);
}

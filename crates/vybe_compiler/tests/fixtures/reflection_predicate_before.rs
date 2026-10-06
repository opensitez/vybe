use vybe_runtime::{Chunk,Op};
use vybe_compiler::primitives::class_slots;
use vybe_compiler::primitives::reflection::{FIELD_TYPES,FIELD_TYPE};
fn emit_import_call(chunks:&mut[Chunk],current:usize,module:&str,name:&str,argc:u8,line:u32){let idx=chunks[current].add_import(module,name);chunks[current].emit_call(idx,argc,line);}
pub fn emit_instanceof(chunks: &mut [Chunk], current: usize, line: u32) {
    let base = chunks[current].local_count;
    chunks[current].alloc_scratch(3);
    let (obj_s, klass_s, types_s) = (base, base + 1, base + 2);
    chunks[current].emit_op_u16(Op::LOCAL_SET, klass_s, line); // [obj]
    chunks[current].emit_op_u16(Op::LOCAL_SET, obj_s, line); // []
    // Type identity lives in class slots, including on Map-backed adapters.
    // Collection indexing would read Map entries instead of named slots.
    let types_key = class_slots::resolve_interned(
        &mut chunks[current],
        &class_slots::ClassSlot::internal(FIELD_TYPES),
        &class_slots::PlainNames,
    );
    let type_key = class_slots::resolve_interned(
        &mut chunks[current],
        &class_slots::ClassSlot::internal(FIELD_TYPE),
        &class_slots::PlainNames,
    );
    chunks[current].emit_op_u16(Op::LOCAL_GET, obj_s, line);
    class_slots::emit_class_get(
        &mut chunks[current],
        class_slots::ObjSource::Stack,
        &types_key,
        class_slots::Dest::Stack,
        line,
    );
    chunks[current].emit_op_u16(Op::LOCAL_SET, types_s, line);
    // if types present -> types.includes(class_name); else obj.__type == class_name
    chunks[current].emit_op_u16(Op::LOCAL_GET, types_s, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, types_s, line);
    emit_import_call(chunks, current, "wasm:js-undefined", "test", 1, line);
    chunks[current].emit_op(Op::I32_OR, line);
    chunks[current].emit_op(Op::I32_EQZ, line); // 1 when __types is present
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, types_s, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, klass_s, line);
    vybe_compiler::primitives::collections::emit_contains(chunks, current, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, obj_s, line);
    class_slots::emit_class_get(
        &mut chunks[current],
        class_slots::ObjSource::Stack,
        &type_key,
        class_slots::Dest::Stack,
        line,
    );
    chunks[current].emit_op_u16(Op::LOCAL_GET, klass_s, line);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    chunks[current].emit_end(line);
}

/// Stack: `[] -> [i32]`. Is the value in `value_slot` an instance of the type
/// spelled `type_name`?
///
/// **The one answer to that question.** Every asker — typed `catch` arms,
/// `instanceof` / `is` / `isinstance` — routes here instead of emitting its own
/// test, so a language cannot get a different answer depending on which
/// construct asked.
///
/// Identity has two representations, and this unions them because neither is
/// complete on its own:
///
/// * the **rtt** — `ref.test` against the declared type index, which resolves
///   declared subtyping and `implements` by index (so it is the half that
///   answers for interfaces), but only for an object that carries an rtt, and
///   only when the most-derived class was the one that allocated it;
/// * **`__types`** — the ancestry chain [`emit_instanceof_chain`] stamps and
///   [`emit_instanceof`] reads, which every frontend writes and which survives
///   a base constructor allocating the instance.
///
/// A bare `ref.test` is therefore not a safe spelling of this question: an
/// object whose rtt is its PARENT's answers "no" to its own class while
/// `__types` says "yes", and the negative rtt suppresses the string chain
/// rather than deferring to it. Testing the rtt FIRST keeps it authoritative
/// wherever it is present and correct; `__types` only ever adds matches.
///
/// Both halves are cross-language: the rtt resolves through the module type
/// index space and `__types` holds canonical names, so a class from one
/// frontend answers for a type declared by another.
pub fn emit_is_instance_of(
    chunks: &mut [Chunk],
    current: usize,
    value_slot: u16,
    type_name: &str,
    line: u32,
) {
    // rtt FIRST, `__types` chain as the fallback.
    //
    // ⚠ The fallback is a KNOWN FORGERY SURFACE — it reads `__type`/`__types`,
    // which are ordinary writable properties, so an object carrying the right
    // string is believed. Deleting it was tried and is NOT yet possible: the
    // VM's own name path (`ref_test_or_declared_name`) does not answer for the
    // untyped values this chain currently carries, so removing it broke real
    // identity (`c instanceof C` → false) and lost the dart typed-catch set.
    //
    // It comes out when nothing is allocated untyped — platform exceptions (27
    // sites) and dynamic literals — and not before. Recorded rather than
    // papered over.
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    let ht = vybe_compiler::primitives::classes::heaptype_for_name(chunks, type_name);
    chunks[current].emit_ref_type_op(Op::REF_TEST, ht, line);
    chunks[current].emit_if_value(line);
    // ⛔ BOTH BRANCHES MUST YIELD THE SAME TYPE, and the conversion belongs
    // AFTER the merge, not inside each arm.
    //
    // The cause of `1`/`0` was the CONVERSION HELPER, not this block: the old
    // code finished with `emit_dyn_to_bool`, which returns an **i32** (see its
    // doc — it is ToBoolean, not the inverse of `emit_i32_to_bool`), so
    // `typeof (c instanceof C)` reported `"number"`.
    //
    // Converting once after the merge is kept because it is cleaner than
    // converting in both arms, and it satisfies the same-type-out-of-both-arms
    // invariant with one instruction sequence instead of two. It was NOT
    // measured to be load-bearing on its own — `emit_i32_to_bool` inside both
    // arms was never tried, so do not cite this position as the fix.
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value_slot, line);
    chunks[current].emit_string_const(type_name, line);
    emit_instanceof(chunks, current, line);
    chunks[current].emit_end(line);
    // ⛔ `emit_i32_to_bool`, NOT `emit_dyn_to_bool`. The two names read like
    // inverses and are not: `emit_i32_to_bool` turns an i32 into a real
    // `Value::Bool`, while `emit_dyn_to_bool` is ToBoolean — it takes ANY
    // dynamic value and returns an **i32**. Converting with the latter is why
    // `c instanceof C` printed `1` and `typeof` reported `"number"`: the value
    // was never a boolean at all. This is the same helper `#x in obj` uses.
    vybe_compiler::primitives::ops::emit_i32_to_bool(&mut chunks[current], line);
}


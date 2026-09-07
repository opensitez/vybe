//! PowerShell operator lowerings that need a RUNTIME type test.

use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

/// `@( … )` — the array subexpression operator. It guarantees an array and
/// flattens one level: `@(1..5)` is five elements, `@($arr)` is `$arr`'s
/// elements, and `@(7)` is one element.
///
/// Whether an operand flattens is a question about its RUNTIME value, so it
/// cannot be decided from the syntax: marking elements as spread drops a scalar
/// (nothing to iterate), and nesting them keeps a collection whole. `concat`
/// answers that for an array, a Set, a Map and a scalar alike.
///
/// ⛔**BUT NOT FOR A STRING.** `ecma:array:concat` deliberately serves TWO
/// callers — spec `Array.prototype.concat` AND spread-element compilation
/// (`[...s]`) — and its own comment says so, so it spreads a string into its
/// code points. That is right for `[...s]` and wrong for both `concat` and
/// `@( … )`: PowerShell's `@("abc")` is ONE element, and it was answering
/// THREE. Measured against `/usr/local/bin/pwsh` — `@("abc").Count` is 1 there
/// and was 3 here, so `@($text)` handed back the text's characters and every
/// `$s.Count` / `$s[0]` on it read a letter.
///
/// The string is therefore appended, never concatenated. Everything else keeps
/// going through `concat`, so an array still flattens and a Set/Map still
/// enumerates — testing `isArray` instead would have taken those two with it.
///
/// ⛔This is not an `ecma` bug to fix in `ecma`: that host fn has a second
/// caller which NEEDS the spreading. The wrong meaning was being asked for.
///
/// Stack: `[a0, …, aN-1]` → `[array]`.
pub fn emit_ensure_array(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let concat = chunk.add_import("ecma:array", "concat");
    let push = chunk.add_import("ecma:array", "push");
    let is_string = chunk.add_import("wasm:js-string", "test");
    let get_method = chunk.add_import("ecma:value", "getMethodForCall");
    let get_prop = chunk.add_import("ecma:object", "get");
    let type_of = chunk.add_import("ecma:value", "typeof");
    let is_array = chunk.add_import("ecma:array", "isArray");
    let cast_bool = chunk.add_import("wasm:js-boolean", "cast");
    let array_from = chunk.add_import("ecma:array", "from");

    // Arguments arrive with the LAST on top, so store them before rebuilding
    // the call in source order.
    let base = chunk.alloc_scratch(argc.max(1) as u16);
    for i in (0..argc as u16).rev() {
        chunk.emit_op_u16(Op::LOCAL_SET, base + i, line);
    }

    // `concat` returns a NEW array while `push` mutates in place, so the
    // accumulator lives in a local rather than on the stack — the two branches
    // cannot leave the same thing there.
    let acc = chunk.alloc_scratch(1);
    chunk.emit_array_new_fixed(0, 0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, acc, line);

    // ⛔EVERY scratch local up front: `alloc_scratch` ALIASES named locals, so
    // allocating inside the loop would hand two iterations the same slot.
    let iter_fn = chunk.alloc_scratch(1);
    let next_fn = chunk.alloc_scratch(1);
    let cur_fn = chunk.alloc_scratch(1);
    let drained = chunk.alloc_scratch(1);
    let guard = chunk.alloc_scratch(1);
    let orig = chunk.alloc_scratch(1);
    let step = chunk.alloc_scratch(1);
    for i in 0..argc as u16 {
        // ⛔THE ORIGINAL IS KEPT, because the `Iterator` slot answers for values
        // this adaptation has no business touching. A native `Set` — which is
        // what `[System.Collections.Generic.HashSet[int]]` is — carries
        // `Symbol.iterator`, so the probe below fired, replaced the Set with a
        // SET ITERATOR OBJECT, and `concat` then appended that object whole:
        // `@($set)` answered ONE element and `foreach ($x in $set)` handed the
        // iterator to `+=`, trapping in `wasm:js-number.toF64`. `concat` already
        // enumerates a Set and a Map correctly, so the replacement is kept ONLY
        // when it produced something better — an array, or a drained enumerator.
        chunk.emit_op_u16(Op::LOCAL_GET, base + i, line);
        chunk.emit_op_u16(Op::LOCAL_SET, orig, line);
        // ⛔`$null` IS ONE ELEMENT, and it is asked NOTHING first. Measured
        // against pwsh 7.6.4: `@($null).Count` is 1 and `$null | ForEach-Object`
        // runs once. Every probe below — `getMethodForCall`, `isArray`,
        // `typeof` — dereferences its operand, so a null reaching them trapped
        // with `null structure reference` rather than being appended.
        chunk.emit_op_u16(Op::LOCAL_GET, orig, line);
        chunk.emit_op(Op::REF_IS_NULL, line);
        chunk.emit_if(line);
        chunk.emit_op_u16(Op::LOCAL_GET, acc, line);
        chunk.emit_op_u16(Op::LOCAL_GET, orig, line);
        chunk.emit_call(push, 2, line);
        chunk.emit_op(Op::DROP, line);
        chunk.emit_else(line);
        // An ENUMERABLE is asked for its enumeration first.
        //
        // `class Col : IEnumerable { [IEnumerator] GetEnumerator() { … } }`
        // reaching a pipeline head is not a scalar and not an array, so
        // `concat` appended the OBJECT whole and `$col | ForEach-Object { $_ * 2 }`
        // answered `NaN` — while the explicit `$col.GetEnumerator() | …`
        // beside it answered `2,4,6`. The class already carries the
        // enumeration: `protocol.rs` maps `GetEnumerator` onto the shared
        // `Iterator` role, whose canonical member name is `iterator`.
        //
        // ⛔THE ROLE IS A SLOT, AND THE SLOT IS NOT ITS SPELLING.
        // `getMethodForCall(col, "iterator")` finds NOTHING, and so does
        // `ClassSlot::internal("iterator")` — `$col.iterator()` is
        // `undefined is not callable`. `classes.rs` installs the member under
        // `canon(source_name)`, so it lives at `getenumerator`; the canonical
        // name only keys `current_class_slot_keys`, which is compile-time.
        // `ClassSlot::Slot(ProtocolSlot::Iterator)` is the stamped runtime
        // key, and asking for the SLOT rather than a spelling is what
        // flexclassplan §2a requires of a language in the first place.
        // ⛔The result REPLACES the operand rather than switching the branch
        // below to `isArray` — `concat` is what flattens a Set and a Map, and
        // this function's own comment records that testing `isArray` instead
        // would take both of those with it.
        chunk.emit_op_u16(Op::LOCAL_GET, base + i, line);
        vybe_compiler::primitives::class_slots::emit_class_get(
            chunk,
            vybe_compiler::primitives::class_slots::ObjSource::Stack,
            &vybe_compiler::primitives::class_slots::resolve(
                &vybe_compiler::primitives::class_slots::ClassSlot::Slot(
                    vybe_ast::ProtocolSlot::Iterator,
                ),
                &vybe_compiler::primitives::class_slots::PlainNames,
            ),
            vybe_compiler::primitives::class_slots::Dest::Local(iter_fn),
            line,
        );

        chunk.emit_op_u16(Op::LOCAL_GET, iter_fn, line);
        chunk.emit_call(type_of, 1, line);
        chunk.emit_string_const("function", line);
        vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
        vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
        chunk.emit_if_value(line);
        // ⛔THE RECEIVER TRAVELS AS ARGUMENT 0. PowerShell declares no
        // `method_receiver`, so the call site supplies one — the same
        // convention `emit_smart_length` uses to call `__get_length`.
        chunk.emit_op_u16(Op::LOCAL_GET, iter_fn, line);
        chunk.emit_op_u16(Op::LOCAL_GET, base + i, line);
        vybe_compiler::primitives::callable::emit_direct_invoke_chunk(chunk, 1, line);
        chunk.emit_else(line);
        chunk.emit_op_u16(Op::LOCAL_GET, base + i, line);
        chunk.emit_end(line);
        chunk.emit_op_u16(Op::LOCAL_SET, base + i, line);

        // A .NET ENUMERATOR IS DRAINED HERE, because it is not an iterator in
        // the sense everything downstream means.
        //
        // `GetEnumerator` may hand back an ARRAY — `$this.Items.GetEnumerator()`
        // — and `concat` below already flattens that. It may equally hand back
        // an object with `MoveNext()` and `get_Current()`, and that object
        // answers NOTHING the shared machinery asks for: the for-of driver
        // wants `next()` returning `{done, value}` per ECMA-262, while
        // `MoveNext` returns a BOOL and the value lives on a separate member.
        // The renames are already in place — `protocol.rs` maps `MoveNext` onto
        // the `Next` role — so the shape is the only thing missing, and this
        // loop IS the adapter.
        //
        // ⛔`get_Current` is read by SPELLING, and only here. No `Current`
        // protocol slot exists to ask instead, and a PowerShell spelling in
        // PowerShell's own emitter is where a spelling belongs — the rule
        // forbids them in SHARED crates.
        // ⛔BOUNDED, and the bound is not a detail. An enumerator is free to be
        // infinite — `[NatEnum] MoveNext() { $this.Val++; return $true }` is in
        // the corpus — so an eager drain MUST stop somewhere or it is a hang,
        // not a wrong answer. `generators.rs` picked 1,000,000 for its own
        // drain; measured here that turned two tests from a fast failure into a
        // 60-SECOND TIMEOUT EACH and took the whole PowerShell gate from 58s to
        // 136s, which every peer pays on every run. Each drained step is a full
        // dynamic invoke — MEASURED at ~1.1ms in a debug build — so 10,000 is
        // 11s per site and still risks the 60s per-test TIMEOUT on a loaded
        // machine. 1,000 is ~1.1s, and it answers every enumerator in the
        // corpus, because what they ask for is `Select-Object -First 3` /
        // `-First 5` off the front of a source that never ends.
        //
        // ⚠ This is a BOUND standing in for laziness, not laziness. A finite
        // enumeration longer than the cap would be silently truncated. The real
        // fix is to hand the pipeline a lazy sequence — the stack-switching
        // generator path — rather than an array; until then the cap is what
        // keeps an infinite source from hanging the process.
        chunk.emit_op_u16(Op::LOCAL_GET, base + i, line);
        vybe_compiler::primitives::class_slots::emit_class_get(
            chunk,
            vybe_compiler::primitives::class_slots::ObjSource::Stack,
            &vybe_compiler::primitives::class_slots::resolve(
                &vybe_compiler::primitives::class_slots::ClassSlot::Slot(
                    vybe_ast::ProtocolSlot::Next,
                ),
                &vybe_compiler::primitives::class_slots::PlainNames,
            ),
            vybe_compiler::primitives::class_slots::Dest::Local(next_fn),
            line,
        );

        // ⛔TWO KEYS, slot THEN spelling — the probe `generators.rs` documents.
        // A .NET enumerator fills the `Next` SLOT. A native ECMA iterator (what
        // `[System.Collections.Generic.HashSet[int]]` hands back) carries the
        // plain `next` PROPERTY and no slot at all, and asking only the slot
        // left it undrained.
        chunk.emit_op_u16(Op::LOCAL_GET, next_fn, line);
        chunk.emit_call(type_of, 1, line);
        chunk.emit_string_const("function", line);
        vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
        vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
        chunk.emit_op(Op::I32_EQZ, line);
        chunk.emit_if(line);
        chunk.emit_op_u16(Op::LOCAL_GET, base + i, line);
        chunk.emit_string_const("next", line);
        chunk.emit_call(get_method, 2, line);
        chunk.emit_op_u16(Op::LOCAL_SET, next_fn, line);
        chunk.emit_end(line);

        chunk.emit_op_u16(Op::LOCAL_GET, next_fn, line);
        chunk.emit_call(type_of, 1, line);
        chunk.emit_string_const("function", line);
        vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
        vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
        chunk.emit_if(line);

        chunk.emit_array_new_fixed(0, 0, line);
        chunk.emit_op_u16(Op::LOCAL_SET, drained, line);
        chunk.emit_i32_const(0, line);
        chunk.emit_op_u16(Op::LOCAL_SET, guard, line);

        let exit = chunk.emit_block(line);
        let (again, _) = chunk.emit_loop_s(line);

        chunk.emit_op_u16(Op::LOCAL_GET, guard, line);
        chunk.emit_i32_const(1_000, line);
        chunk.emit_op(Op::I32_GE_S, line);
        chunk.emit_br_if(1, line);

        chunk.emit_op_u16(Op::LOCAL_GET, next_fn, line);
        chunk.emit_op_u16(Op::LOCAL_GET, base + i, line);
        vybe_compiler::primitives::callable::emit_direct_invoke_chunk(chunk, 1, line);
        chunk.emit_op_u16(Op::LOCAL_SET, step, line);

        // ⛔THE STEP HAS TWO SHAPES and the RESULT tells them apart. ECMA-262
        // says `next()` returns an OBJECT carrying `done`/`value`; .NET's
        // `MoveNext()` returns a BOOL and leaves the value on `Current`. Both
        // reach this loop, so the test is on what came back, not on which key
        // found the method.
        chunk.emit_op_u16(Op::LOCAL_GET, step, line);
        chunk.emit_call(type_of, 1, line);
        chunk.emit_string_const("object", line);
        vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
        vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
        chunk.emit_if(line);

        chunk.emit_op_u16(Op::LOCAL_GET, step, line);
        chunk.emit_string_const("done", line);
        chunk.emit_call(get_prop, 2, line);
        vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
        chunk.emit_br_if(2, line);
        chunk.emit_op_u16(Op::LOCAL_GET, drained, line);
        chunk.emit_op_u16(Op::LOCAL_GET, step, line);
        chunk.emit_string_const("value", line);
        chunk.emit_call(get_prop, 2, line);
        chunk.emit_call(push, 2, line);
        chunk.emit_op(Op::DROP, line);

        chunk.emit_else(line);

        chunk.emit_op_u16(Op::LOCAL_GET, step, line);
        vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
        chunk.emit_op(Op::I32_EQZ, line);
        chunk.emit_br_if(2, line);

        // ⛔`get_Current` is read by SPELLING, and only here. No `Current`
        // protocol slot exists to ask instead, and a PowerShell spelling in
        // PowerShell's own emitter is where a spelling belongs.
        chunk.emit_op_u16(Op::LOCAL_GET, base + i, line);
        chunk.emit_string_const("get_current", line);
        chunk.emit_call(get_method, 2, line);
        chunk.emit_op_u16(Op::LOCAL_SET, cur_fn, line);

        chunk.emit_op_u16(Op::LOCAL_GET, drained, line);
        chunk.emit_op_u16(Op::LOCAL_GET, cur_fn, line);
        chunk.emit_call(type_of, 1, line);
        chunk.emit_string_const("function", line);
        vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
        vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
        chunk.emit_if_value(line);
        chunk.emit_op_u16(Op::LOCAL_GET, cur_fn, line);
        chunk.emit_op_u16(Op::LOCAL_GET, base + i, line);
        vybe_compiler::primitives::callable::emit_direct_invoke_chunk(chunk, 1, line);
        chunk.emit_else(line);
        chunk.emit_op_u16(Op::LOCAL_GET, base + i, line);
        chunk.emit_string_const("current", line);
        chunk.emit_call(get_prop, 2, line);
        chunk.emit_end(line);
        chunk.emit_call(push, 2, line);
        chunk.emit_op(Op::DROP, line);

        chunk.emit_end(line);

        chunk.emit_op_u16(Op::LOCAL_GET, guard, line);
        chunk.emit_i32_const(1, line);
        chunk.emit_op(Op::I32_ADD, line);
        chunk.emit_op_u16(Op::LOCAL_SET, guard, line);

        chunk.emit_br(0, line);
        chunk.emit_end(line);
        chunk.patch_loop(again);
        chunk.emit_end(line);
        chunk.patch_block(exit);

        chunk.emit_op_u16(Op::LOCAL_GET, drained, line);
        chunk.emit_op_u16(Op::LOCAL_SET, base + i, line);
        chunk.emit_end(line);

        // Still not an array: either the slot answered with a plain ECMA
        // iterable, or this value never entered the adaptation at all.
        //
        // ⛔A `HashSet` IS A NATIVE `Set` — `$set.ToString()` is `[object Set]`
        // — and `concat` does NOT flatten one, measured: `@($set).Count` is 1.
        // That was invisible until `foreach` started routing its source through
        // here, because the shared for-of iterates a Set natively and `@()`
        // over one is in no test. `Array.from` is the primitive for exactly
        // this and covers a Set, a Map and an iterator object alike.
        //
        // ⛔Gated on the value HAVING a `values` method, not on `Array.from`
        // succeeding: `Array.from` answers `[]` for a NON-iterable object, and
        // a PowerShell `@{ a = 1 }` is a bare object that must stay ONE element.
        chunk.emit_op_u16(Op::LOCAL_GET, base + i, line);
        chunk.emit_call(is_array, 1, line);
        chunk.emit_call(cast_bool, 1, line);
        chunk.emit_op(Op::I32_EQZ, line);
        chunk.emit_if(line);

        chunk.emit_op_u16(Op::LOCAL_GET, base + i, line);
        chunk.emit_string_const("values", line);
        chunk.emit_call(get_method, 2, line);
        chunk.emit_call(type_of, 1, line);
        chunk.emit_string_const("function", line);
        vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
        vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
        chunk.emit_if(line);
        chunk.emit_op_u16(Op::LOCAL_GET, base + i, line);
        chunk.emit_call(array_from, 1, line);
        chunk.emit_op_u16(Op::LOCAL_SET, base + i, line);
        chunk.emit_else(line);
        chunk.emit_op_u16(Op::LOCAL_GET, orig, line);
        chunk.emit_op_u16(Op::LOCAL_SET, base + i, line);
        chunk.emit_end(line);

        chunk.emit_end(line);

        chunk.emit_op_u16(Op::LOCAL_GET, base + i, line);
        chunk.emit_call(is_string, 1, line);
        chunk.emit_if(line);

        chunk.emit_op_u16(Op::LOCAL_GET, acc, line);
        chunk.emit_op_u16(Op::LOCAL_GET, base + i, line);
        chunk.emit_call(push, 2, line);
        chunk.emit_op(Op::DROP, line);

        chunk.emit_else(line);

        chunk.emit_op_u16(Op::LOCAL_GET, acc, line);
        chunk.emit_op_u16(Op::LOCAL_GET, base + i, line);
        chunk.emit_call(concat, 2, line);
        chunk.emit_op_u16(Op::LOCAL_SET, acc, line);

        chunk.emit_end(line);

        // Closes the null guard opened at the top of the loop body.
        chunk.emit_end(line);
    }

    chunk.emit_op_u16(Op::LOCAL_GET, acc, line);
}

/// `$obj.PSObject.Properties.Add($descriptor)` — attach an ETS member.
///
/// Two kinds reach here, both pure data:
///
/// * `note` — the payload IS the value. `$o.Status = "Active"`.
/// * `alias` — the payload NAMES another member, so the value is read off the
///   target. ⚠ Read ONCE, here: a real `PSAliasProperty` re-reads its target on
///   every access. Nothing in the corpus writes through an alias or mutates the
///   aliased member afterwards, and a stored value is a truthful snapshot where
///   a fake getter would not be.
///
/// A descriptor of any other kind is left alone rather than half-applied.
///
/// Stack: `[target, descriptor]` -> `[null]`.
/// `.Add(…)` — one spelling, two collections, told apart by ARITY.
///
/// A list takes one element (`$list.Add($x)`); a hashtable takes a key AND a
/// value (`$h.Add($k, $v)`), and its storage is object properties rather than
/// array slots. The arity is known here, so nothing has to guess from the name.
///
/// ⛔ This is the fallback for an UNTYPED receiver only. A `Dictionary` names a
/// registered type and resolves through the namespace tree to `ecma:map`
/// before this row is consulted, which is what keeps the two storages apart —
/// the previous walker rewrite keyed on spelling and arity alone and sent a
/// `Dictionary`'s `Add` to object properties while its `Count` and `Keys` read
/// the map.
///
/// Stack: `[recv, value]` → `[recv']`, or `[recv, key, value]` → `[…]`.
pub fn emit_collection_add(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    if argc == 3 {
        let idx = chunks[current].add_import("ecma:object", "set");
        chunks[current].emit_call(idx, 3, line);
    } else {
        vybe_compiler::primitives::collections::emit_push(chunks, current, line);
    }
}

pub fn emit_prop_add(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let desc = chunk.alloc_scratch(1);
    let target = chunk.alloc_scratch(1);
    let member = chunk.alloc_scratch(1);
    let payload = chunk.alloc_scratch(1);

    let obj_get = chunk.add_import("ecma:object", "get");
    let obj_set = chunk.add_import("ecma:object", "set");

    chunk.emit_op_u16(Op::LOCAL_SET, desc, line);
    chunk.emit_op_u16(Op::LOCAL_SET, target, line);

    chunk.emit_op_u16(Op::LOCAL_GET, desc, line);
    chunk.emit_string_const("Name", line);
    chunk.emit_call(obj_get, 2, line);
    chunk.emit_op_u16(Op::LOCAL_SET, member, line);

    // ⛔The value is computed into a LOCAL first. Leaving `[target, member]`
    // underneath an `if_value` block and joining the three on the stack
    // afterwards produced `NaN` at every call site — the operands never met.
    chunk.emit_op_u16(Op::LOCAL_GET, desc, line);
    chunk.emit_string_const("Value", line);
    chunk.emit_call(obj_get, 2, line);
    chunk.emit_op_u16(Op::LOCAL_SET, payload, line);

    chunk.emit_op_u16(Op::LOCAL_GET, desc, line);
    chunk.emit_string_const("__ps_kind", line);
    chunk.emit_call(obj_get, 2, line);
    chunk.emit_string_const("alias", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::LOCAL_GET, target, line);
    chunk.emit_op_u16(Op::LOCAL_GET, payload, line);
    chunk.emit_call(obj_get, 2, line);
    chunk.emit_op_u16(Op::LOCAL_SET, payload, line);
    chunk.emit_end(line);

    chunk.emit_op_u16(Op::LOCAL_GET, target, line);
    chunk.emit_op_u16(Op::LOCAL_GET, member, line);
    chunk.emit_op_u16(Op::LOCAL_GET, payload, line);
    chunk.emit_call(obj_set, 3, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_ref_null(0x6e, line);
}

/// `$obj.PSObject` — the EXTENDED TYPE SYSTEM view of a value.
///
/// PowerShell wraps every value in a `PSObject` carrying `TypeNames`,
/// `Properties` and `Members` alongside the value's own members. It has to be
/// the SAME object every time, because the corpus mutates through it and reads
/// back: `$o.psobject.TypeNames.Insert(0, "T")` then
/// `$o.psobject.TypeNames[0]`. So this is get-or-CREATE, stashed on the target
/// under a private key, not a fresh view per access.
///
/// ⛔`TypeNames` ARRIVES AS AN ARGUMENT rather than being built here. It must
/// support `Insert`/`Add`/`Contains`/`Clear`/`Count`/`[0]`, and a bare array
/// answers only some of those — `@("a").Insert(0, "b")` is
/// `undefined is not callable`, because `Insert` is a leaf on the dotnet TREE
/// and declaring a `[value_methods] Insert` row would SHADOW that leaf and
/// break `List[T].Insert` for everyone. The walker therefore hands in a real
/// `List[string]`, built the ordinary way, and this only stores it.
///
/// Stack: `[target, fresh_type_names]` -> `[psobject]`.
pub fn emit_psobject(chunks: &mut [Chunk], current: usize, line: u32) {
    let names = chunks[current].alloc_scratch(1);
    let target = chunks[current].alloc_scratch(1);
    let view = chunks[current].alloc_scratch(1);

    let obj_get = chunks[current].add_import("ecma:object", "get");
    let obj_set = chunks[current].add_import("ecma:object", "set");
    let type_of = chunks[current].add_import("ecma:value", "typeof");

    chunks[current].emit_op_u16(Op::LOCAL_SET, names, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, target, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, target, line);
    chunks[current].emit_string_const("__ps_psobject", line);
    chunks[current].emit_call(obj_get, 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, view, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, view, line);
    chunks[current].emit_call(type_of, 1, line);
    chunks[current].emit_string_const("object", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);

    vybe_compiler::primitives::dict::emit_new(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, view, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, view, line);
    chunks[current].emit_string_const("TypeNames", line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, names, line);
    chunks[current].emit_call(obj_set, 3, line);
    chunks[current].emit_op(Op::DROP, line);

    // The view has to know what it is a view OF: `Properties.Add` writes to the
    // TARGET, not to the view.
    chunks[current].emit_op_u16(Op::LOCAL_GET, view, line);
    chunks[current].emit_string_const("__ps_target", line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, target, line);
    chunks[current].emit_call(obj_set, 3, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, target, line);
    chunks[current].emit_string_const("__ps_psobject", line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, view, line);
    chunks[current].emit_call(obj_set, 3, line);
    chunks[current].emit_op(Op::DROP, line);

    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, view, line);
}

/// PowerShell's `*`, which is the same LEFT-OPERAND rule as `+` — and is the
/// reason `+` needed one in the first place.
///
/// | `$a` | `$a * 3` |
/// |------|----------|
/// | array | the elements REPEATED — `@(1,2) * 3` is six elements |
/// | string | the text repeated — `"ab" * 3` is `"ababab"` |
/// | number | arithmetic |
///
/// Verified against `/usr/local/bin/pwsh` 7.6.4: `(@(1,2) * 3).Count` is 6 and
/// `"ab" * 3` is `ababab` there, while both answered `NaN`/0 here — `F64_MUL`
/// coerces an array or a string operand to a number, so a repetition became a
/// silent nothing rather than an error.
///
/// ⛔The COUNT is read as an i32 and the loop is a real one: `concat` appends
/// one operand per call, so a repetition is n appends. Reusing `emit_add`'s
/// shape would only have appended once.
///
/// Stack: `[a, b]` -> `[result]`.
pub fn emit_multiply(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let b_slot = chunk.alloc_scratch(1);
    let a_slot = chunk.alloc_scratch(1);
    let acc = chunk.alloc_scratch(1);
    let count = chunk.alloc_scratch(1);

    let is_array = chunk.add_import("ecma:array", "isArray");
    let cast_bool = chunk.add_import("wasm:js-boolean", "cast");
    let concat = chunk.add_import("ecma:array", "concat");
    let is_string = chunk.add_import("wasm:js-string", "test");
    let repeat = chunk.add_import("ecma:string", "repeat");

    chunk.emit_op_u16(Op::LOCAL_SET, b_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, a_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, a_slot, line);
    chunk.emit_call(is_array, 1, line);
    chunk.emit_call(cast_bool, 1, line);
    chunk.emit_if_value(line);

    // Array on the left: n appends of the whole array.
    chunk.emit_array_new_fixed(0, 0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, acc, line);
    // ⛔The counter stays a DYNAMIC value. `$n` arrives boxed and there is no
    // `emit_dyn_to_i32`; the shared drain in `generators.rs` counts the same
    // way, with `emit_dyn_lt` / `emit_dyn_add`.
    chunk.emit_op_u16(Op::LOCAL_GET, b_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, count, line);

    let done = chunk.emit_block(line);
    let (again, _) = chunk.emit_loop_s(line);
    chunk.emit_i32_const(0, line);
    chunk.emit_op_u16(Op::LOCAL_GET, count, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_br_if(1, line);
    chunk.emit_op_u16(Op::LOCAL_GET, acc, line);
    chunk.emit_op_u16(Op::LOCAL_GET, a_slot, line);
    chunk.emit_call(concat, 2, line);
    chunk.emit_op_u16(Op::LOCAL_SET, acc, line);
    chunk.emit_op_u16(Op::LOCAL_GET, count, line);
    chunk.emit_i32_const(-1, line);
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_SET, count, line);
    chunk.emit_br(0, line);
    chunk.emit_end(line);
    chunk.patch_loop(again);
    chunk.emit_end(line);
    chunk.patch_block(done);
    chunk.emit_op_u16(Op::LOCAL_GET, acc, line);

    chunk.emit_else(line);

    chunk.emit_op_u16(Op::LOCAL_GET, a_slot, line);
    chunk.emit_call(is_string, 1, line);
    chunk.emit_if_value(line);

    // String on the left: `String.prototype.repeat`, which is this operation
    // exactly — including throwing on a negative count, as PowerShell does.
    chunk.emit_op_u16(Op::LOCAL_GET, a_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, b_slot, line);
    chunk.emit_call(repeat, 2, line);

    chunk.emit_else(line);

    // Number on the left: arithmetic — but ⛔NOT a bare `F64_MUL`. `*` still has
    // to reach a user `operator *`, which is what `emit_divide` does for `/`
    // through the same helper. Multiplying an `Int128`/`BigInteger` numerically
    // would answer nonsense, and those carry their own operator.
    // ⛔`emit_rich_arithmetic` takes the operands as LOCAL SLOTS, not from the
    // stack — pushing them first leaves two values behind and unbalances the
    // enclosing `if`.
    vybe_compiler::primitives::expressions::emit_rich_arithmetic(
        chunk,
        a_slot,
        b_slot,
        &vybe_ast::protocol_slot_key(vybe_ast::ProtocolSlot::Mul),
        |c: &mut Chunk, l: u32| c.emit_op(Op::F64_MUL, l),
        line,
    );

    chunk.emit_end(line);
    chunk.emit_end(line);
}

/// PowerShell's `+`, which is three operations chosen by the LEFT operand:
///
/// | `$a` | `$a + $b` |
/// |------|-----------|
/// | array | `$b`'s elements appended — `@(1,2) + 3` is `@(1,2,3)` |
/// | string | concatenation — `'5' + 5` is `'55'` |
/// | number | arithmetic — `5 + '5'` is `10` |
///
/// No shared primitive answers this. `F64_ADD` and `dynamic_add` both coerce an
/// array operand to a number (`@(1,2) + 3` became `NaN`, and `$a += $x` — the
/// idiomatic way to grow an array in PowerShell — trapped in
/// `wasm:js-number.toF64`). `[builtin_slots.array] add` is not a way out: the
/// shared `compile_binop` consults it for EVERY `+`, so binding
/// `collections.concat` there made `10 + 5` evaluate to `5`.
///
/// The left-operand rule is why this cannot be `dynamic_add` either: that
/// concatenates whenever EITHER side is a string, so it answers `'55'` for
/// `5 + '5'` where PowerShell answers `10`.
///
/// Stack: `[a, b]` → `[result]`.
pub fn emit_add(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let b_slot = chunk.alloc_scratch(1);
    let a_slot = chunk.alloc_scratch(1);

    let is_array = chunk.add_import("ecma:array", "isArray");
    let cast_bool = chunk.add_import("wasm:js-boolean", "cast");
    let concat = chunk.add_import("ecma:array", "concat");
    let is_string = chunk.add_import("wasm:js-string", "test");

    chunk.emit_op_u16(Op::LOCAL_SET, b_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, a_slot, line);

    // `$null` on the left is the IDENTITY, whatever the right operand is.
    // Measured against pwsh 7.6.4: `$null + 1` is `1`, `$null + 'x'` is `x`,
    // `$null + @(1,2)` is the two-element array and `$null + $null` is `$null`
    // — never a number. Falling through to `F64_ADD` answered `0`, `NaN` and
    // `0`, and an UNASSIGNED variable made it worse: it reads as `$null` in
    // PowerShell but arrives here as `undefined`, whose `as_f64` is `NaN`, so
    // the accumulator idiom `$sum += $n` produced `NaN` from its first item.
    let nullish = |chunk: &mut Chunk, slot: u16| {
        chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
        chunk.emit_op(Op::REF_IS_NULL, line);
        chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
        let idx = chunk.add_import("wasm:js-undefined", "test");
        chunk.emit_call(idx, 1, line);
        chunk.emit_op(Op::I32_OR, line);
    };

    nullish(chunk, a_slot);
    chunk.emit_if_value(line);
    chunk.emit_op_u16(Op::LOCAL_GET, b_slot, line);
    chunk.emit_else(line);

    // `isArray` answers with a boxed boolean; the branch needs an i32.
    chunk.emit_op_u16(Op::LOCAL_GET, a_slot, line);
    chunk.emit_call(is_array, 1, line);
    chunk.emit_call(cast_bool, 1, line);
    chunk.emit_if_value(line);

    // Array on the left: append. `concat` flattens an array operand and
    // appends a scalar one, which is exactly what `+` means here.
    chunk.emit_op_u16(Op::LOCAL_GET, a_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, b_slot, line);
    chunk.emit_call(concat, 2, line);

    chunk.emit_else(line);

    // String on the left: concatenate, coercing the right. Note this tests only
    // the LEFT operand — `emit_dyn_add` would concatenate whenever EITHER side
    // is a string and so answer `'55'` for `5 + '5'`, where PowerShell answers
    // `10` because the left operand is a number.
    chunk.emit_op_u16(Op::LOCAL_GET, a_slot, line);
    chunk.emit_call(is_string, 1, line);
    chunk.emit_if_value(line);

    // `emit_dyn_add` concatenates whenever either operand is a string, and the
    // left one is — so here it is exactly right, coercions included.
    //
    // ⛔EXCEPT FOR A NULL RIGHT OPERAND, which it renders as `0`: `'x' + $null`
    // answered `x0` where pwsh answers `x`. PowerShell appends nothing, so the
    // empty string is substituted before the concatenation sees it.
    chunk.emit_op_u16(Op::LOCAL_GET, a_slot, line);
    nullish(chunk, b_slot);
    chunk.emit_if_value(line);
    chunk.emit_string_const("", line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, b_slot, line);
    chunk.emit_end(line);
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);

    chunk.emit_else(line);

    // Number on the left: arithmetic. `F64_ADD` coerces BOTH operands through
    // `Value::as_f64`, which is what makes `5 + '5'` equal `10`.
    chunk.emit_op_u16(Op::LOCAL_GET, a_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, b_slot, line);
    chunk.emit_op(Op::F64_ADD, line);

    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_end(line);
}

/// `$a.CompareTo($b)` — `System.Object`'s comparison, which every PowerShell
/// value has: `(5).CompareTo(3)` is `1`, `('b').CompareTo('a')` is `1`, and
/// equal values give `0`.
///
/// ⛔**NOT `object.compare`.** That primitive's stack is `[a, b, comparator]`
/// and it INVOKES the comparator unconditionally — there is no null-comparator
/// branch, despite the name. Binding a two-operand `[value_methods]` row to it
/// read a garbage third operand (both examples above answered `-1`), and
/// supplying an explicit `ref.null` only moved the failure to
/// `null is not callable`. It is the comparator-SUPPLIED form of the operation,
/// not the default one.
///
/// The default ordering is built from the shared dynamic comparison ops
/// instead, so `<` and `>` mean here exactly what they mean everywhere else in
/// the language — including for strings, where they are ordinal.
///
/// Stack: `[a, b]` → `[i32]`.
pub fn emit_compare_to(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let b = chunk.alloc_scratch(2);
    let a = b + 1;
    chunk.emit_op_u16(Op::LOCAL_SET, b, line);
    chunk.emit_op_u16(Op::LOCAL_SET, a, line);

    chunk.emit_op_u16(Op::LOCAL_GET, a, line);
    chunk.emit_op_u16(Op::LOCAL_GET, b, line);
    vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    chunk.emit_i32_const(-1, line);
    chunk.emit_else(line);

    chunk.emit_op_u16(Op::LOCAL_GET, a, line);
    chunk.emit_op_u16(Op::LOCAL_GET, b, line);
    vybe_compiler::primitives::ops::emit_dyn_gt(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    chunk.emit_i32_const(1, line);
    chunk.emit_else(line);
    chunk.emit_i32_const(0, line);
    chunk.emit_end(line);

    chunk.emit_end(line);
}

/// `[char]$x` — the cast to `System.Char`, which is TWO conversions chosen by
/// the operand's runtime type:
///
/// | `$x` | `[char]$x` |
/// |------|------------|
/// | number | the code point — `[char]65` is `A` |
/// | string | the character itself — `[char]'H'` is `H` |
///
/// The slot was bound straight to `strings.from_char_code`, which is only the
/// first row. `[char]'H'` therefore asked for the code point OF A STRING, got a
/// non-number, and rendered as a SPACE — so `[char]::IsUpper([char]'H')` was
/// false and every character-classification test failed on the CAST rather than
/// on the classification it was written to check.
///
/// ⛔The two directions are already declared as a matched pair
/// (`[builtin_slots.char] int` / `[builtin_slots.int] char`) precisely so a
/// cast cannot resolve one way and not the other. That pairing held; what was
/// missing is that ONE of the directions has two source types.
///
/// Stack: `[value]` → `[char]`.
pub fn emit_to_char(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let value = chunk.alloc_scratch(1);
    let is_string = chunk.add_import("wasm:js-string", "test");
    let from_code = chunk.add_import("wasm:js-string", "fromCharCode");

    chunk.emit_op_u16(Op::LOCAL_SET, value, line);

    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    chunk.emit_call(is_string, 1, line);
    chunk.emit_if_value(line);

    // Already a character: a PowerShell `[char]` IS a one-character string.
    chunk.emit_op_u16(Op::LOCAL_GET, value, line);

    chunk.emit_else(line);

    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    chunk.emit_call(from_code, 1, line);

    chunk.emit_end(line);
}

/// `[int]$x` — PowerShell's integer cast, which is THREE conversions:
///
/// | `$x` | `[int]$x` |
/// |------|-----------|
/// | number | banker's rounding — `[int]3.5` is **4**, `[int]2.5` is **2** |
/// | `"65"` | 65 — a numeric parse |
/// | `'A'`  | 65 — the CODE POINT |
///
/// PowerShell tells the last two apart by TYPE: `[int][char]'A'` is the code
/// point and `[int]'A'` on a genuine string is an error. We represent a
/// `[char]` as a one-character string, so the type that would have decided is
/// gone by the time the cast runs and the question has to be asked of the
/// VALUE instead.
///
/// ⛔`char_arithmetic_offset` looked like a REGRESSION from the `[char]` fix
/// and was not: `[char]'A'` used to yield a SPACE and `[int]` of it yielded
/// NaN, so `[char]([int]$ch + 1)` and `[char]'B'` were BOTH a space and the
/// test passed by comparing two wrong answers. Fixing the cast made it fail
/// honestly. Gate a cast change on the whole char suite, never on the one test
/// that names it.
///
/// ⛔The midpoint policy is the shared `MidpointPolicy::HalfEven` primitive,
/// not a hand-rolled `floor(x + 0.5)`. The old lowering answered `3` for
/// `[int]3.5` where PowerShell answers `4`.
///
/// Stack: `[value]` → `[number]`.
pub fn emit_to_int(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let value = chunk.alloc_scratch(2);
    let parsed = value + 1;
    let is_string = chunk.add_import("wasm:js-string", "test");
    let to_number = chunk.add_import("ecma:number", "Number");
    let is_nan = chunk.add_import("ecma:number", "isNaN");
    let char_code = chunk.add_import("wasm:js-string", "charCodeAt");

    chunk.emit_op_u16(Op::LOCAL_SET, value, line);

    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    chunk.emit_call(is_string, 1, line);
    chunk.emit_if_value(line);

    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    chunk.emit_call(to_number, 1, line);
    chunk.emit_op_u16(Op::LOCAL_SET, parsed, line);

    chunk.emit_op_u16(Op::LOCAL_GET, parsed, line);
    chunk.emit_call(is_nan, 1, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    // Not numeric — read it as a character.
    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    chunk.emit_i32_const(0, line);
    chunk.emit_call(char_code, 2, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, parsed, line);
    vybe_compiler::primitives::math::emit_round(
        chunk,
        vybe_ast::MidpointPolicy::HalfEven,
        line,
    );
    chunk.emit_end(line);

    chunk.emit_else(line);

    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    chunk.emit_call(to_number, 1, line);
    vybe_compiler::primitives::math::emit_round(
        chunk,
        vybe_ast::MidpointPolicy::HalfEven,
        line,
    );

    chunk.emit_end(line);
}

/// `$obj.$name` — a computed member read, where the member is the VALUE held by
/// `$name`.
///
/// ⛔`ExprKind::Index` is the obvious node and the wrong one. Indexing an
/// OBJECT with a non-literal key falls through to ARRAY indexing, so `$h['A']`
/// worked (a literal key compiles to a member read) while `$h[$k]` answered
/// EMPTY for the same key. The shared `[builtin_slots.map] get_item` binding
/// exists but never fires here, because a PowerShell `@{…}` walks to
/// `ExprKind::Object` and is not classified as a Map.
///
/// PowerShell member lookup is CASE-INSENSITIVE — `$obj.UNIQUECODE` and
/// `$obj.uniquecode` both find `UniqueCode` — and the compiler already folds
/// declared member names down, so the key is folded here to match. A static
/// `$obj.UNIQUECODE` gets that fold at compile time; a computed one has to ask
/// for it at runtime.
///
/// Stack: `[obj, key]` → `[value_or_null]`.
pub fn emit_member_dyn(chunks: &mut [Chunk], current: usize, line: u32) {
    vybe_compiler::primitives::strings::emit_to_lower(&mut chunks[current], line);
    vybe_compiler::primitives::dict::emit_get_dynamic(chunks, current, line);
}

/// PowerShell `.Count`: null is 0, arrays are their item count, non-empty maps
/// are their map size, and ordinary scalars are one. This is intentionally not
/// shared `collections.length`, whose object fallback is key-count and is right
/// for other languages' `len`/`size` questions but wrong for PowerShell scalar
/// pipeline records.
pub fn emit_count(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let value = chunk.alloc_scratch(2);
    let is_array = chunk.add_import("ecma:array", "isArray");
    let cast_bool = chunk.add_import("wasm:js-boolean", "cast");
    let arr_len = chunk.add_import("ecma:array", "length");
    let map_size = chunk.add_import("ecma:map", "size");
    let to_number = chunk.add_import("ecma:number", "Number");
    let is_undefined = chunk.add_import("wasm:js-undefined", "test");
    let type_of = chunk.add_import("ecma:value", "typeof");

    chunk.emit_op_u16(Op::LOCAL_SET, value, line);

    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if_value(line);
    chunk.emit_f64_const(0.0, line);
    chunk.emit_else(line);

    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    chunk.emit_call(is_array, 1, line);
    chunk.emit_call(cast_bool, 1, line);
    chunk.emit_if_value(line);
    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    chunk.emit_call(arr_len, 1, line);
    chunk.emit_call(to_number, 1, line);
    chunk.emit_else(line);

    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    chunk.emit_call(type_of, 1, line);
    chunk.emit_string_const("object", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);

    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    chunk.emit_string_const("count", line);
    chunk.emit_op(Op::ARRAY_GET, line);
    chunk.emit_op_u16(Op::LOCAL_SET, value + 1, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value + 1, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value + 1, line);
    chunk.emit_call(is_undefined, 1, line);
    chunk.emit_op(Op::I32_OR, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_if_value(line);
    chunk.emit_op_u16(Op::LOCAL_GET, value + 1, line);
    chunk.emit_else(line);

    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    chunk.emit_call(map_size, 1, line);
    chunk.emit_op_u16(Op::LOCAL_SET, value + 1, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value + 1, line);
    chunk.emit_i32_const(0, line);
    chunk.emit_op(Op::I32_NE, line);
    chunk.emit_if_value(line);
    chunk.emit_op_u16(Op::LOCAL_GET, value + 1, line);
    chunk.emit_call(to_number, 1, line);
    chunk.emit_else(line);
    chunk.emit_f64_const(1.0, line);
    chunk.emit_end(line);

    chunk.emit_end(line);
    chunk.emit_else(line);
    chunk.emit_f64_const(1.0, line);
    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_end(line);
}

/// `$h[$k] = $v` / `$obj.$prop = $v` — a computed WRITE, dispatched on the
/// receiver's runtime type.
///
/// ⛔The mirror of `emit_member_dyn`, and it failed the same way: an `Index`
/// assignment resolves its key at COMPILE time, so `$h['A'] = 1` wrote
/// correctly while `$h[$k] = 2` wrote NOWHERE — silently, with no error and no
/// trap. Reading it back gave empty, which reads as "the write was lost"
/// rather than "the write never happened".
///
/// An array keeps `ecma:array:set` so `$a[$i] = v` is untouched; anything else
/// is a property write under the folded key, matching the read path's fold.
///
/// Stack: `[obj, key, value]` → `[]`.
pub fn emit_index_set(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let value = chunk.alloc_scratch(3);
    let key = value + 1;
    let obj = value + 2;
    let is_array = chunk.add_import("ecma:array", "isArray");
    let cast_bool = chunk.add_import("wasm:js-boolean", "cast");
    let array_set = chunk.add_import("ecma:array", "set");

    chunk.emit_op_u16(Op::LOCAL_SET, value, line);
    chunk.emit_op_u16(Op::LOCAL_SET, key, line);
    chunk.emit_op_u16(Op::LOCAL_SET, obj, line);

    chunk.emit_op_u16(Op::LOCAL_GET, obj, line);
    chunk.emit_call(is_array, 1, line);
    chunk.emit_call(cast_bool, 1, line);
    chunk.emit_if(line);

    chunk.emit_op_u16(Op::LOCAL_GET, obj, line);
    chunk.emit_op_u16(Op::LOCAL_GET, key, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    chunk.emit_call(array_set, 3, line);
    chunk.emit_op(Op::DROP, line);

    chunk.emit_else(line);

    chunk.emit_op_u16(Op::LOCAL_GET, obj, line);
    chunk.emit_op_u16(Op::LOCAL_GET, key, line);
    vybe_compiler::primitives::strings::emit_to_lower(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    vybe_compiler::primitives::dict::emit_set_dynamic(chunks, current, line);

    chunks[current].emit_end(line);
}

/// `$h[$k]` — a computed index READ, dispatched on the receiver's runtime type.
/// The read half of `emit_index_set`; see that comment for why a literal key
/// worked and a variable one did not.
///
/// Stack: `[obj, key]` → `[value]`.
pub fn emit_index_get(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let key = chunk.alloc_scratch(4);
    let obj = key + 1;
    let raw = key + 2;
    let fallback = key + 3;
    let is_array = chunk.add_import("ecma:array", "isArray");
    let cast_bool = chunk.add_import("wasm:js-boolean", "cast");
    let array_get = chunk.add_import("ecma:array", "get");
    let object_get = chunk.add_import("ecma:object", "get");
    let is_undefined = chunk.add_import("wasm:js-undefined", "test");

    chunk.emit_op_u16(Op::LOCAL_SET, key, line);
    chunk.emit_op_u16(Op::LOCAL_SET, obj, line);

    chunk.emit_op_u16(Op::LOCAL_GET, obj, line);
    chunk.emit_call(is_array, 1, line);
    chunk.emit_call(cast_bool, 1, line);
    chunk.emit_if_value(line);

    chunk.emit_op_u16(Op::LOCAL_GET, obj, line);
    chunk.emit_op_u16(Op::LOCAL_GET, key, line);
    chunk.emit_call(object_get, 2, line);
    chunk.emit_op_u16(Op::LOCAL_SET, raw, line);
    chunk.emit_op_u16(Op::LOCAL_GET, raw, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_op_u16(Op::LOCAL_GET, raw, line);
    chunk.emit_call(is_undefined, 1, line);
    chunk.emit_op(Op::I32_OR, line);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::LOCAL_GET, obj, line);
    chunk.emit_op_u16(Op::LOCAL_GET, key, line);
    chunk.emit_call(array_get, 2, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, raw, line);
    chunk.emit_end(line);

    chunk.emit_else(line);

    chunk.emit_op_u16(Op::LOCAL_GET, obj, line);
    chunk.emit_op_u16(Op::LOCAL_GET, key, line);
    vybe_compiler::primitives::collections::emit_get(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, raw, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, raw, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, obj, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key, line);
    vybe_compiler::primitives::strings::emit_to_lower(&mut chunks[current], line);
    vybe_compiler::primitives::dict::emit_get_dynamic(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, fallback, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, fallback, line);
    chunks[current].emit_op(Op::REF_IS_NULL, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, fallback, line);
    chunks[current].emit_call(is_undefined, 1, line);
    chunks[current].emit_op(Op::I32_OR, line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, key, line);
    chunks[current].emit_i32_const(0, line);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, obj, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, fallback, line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, fallback, line);
    chunks[current].emit_end(line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, raw, line);
    chunks[current].emit_end(line);

    chunks[current].emit_end(line);
}

/// `… | Out-Null` — evaluate the pipeline, discard its value.
///
/// Stack: `[value]` → `[null]`.
pub fn emit_out_null(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    chunk.emit_op(Op::DROP, line);
    chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
}

/// One argument of an expanded splat: `$h` by NAME when it is a hashtable,
/// `$arr` by INDEX when it is an array.
///
/// ⛔PowerShell decides between the two by the splat's RUNTIME shape, not by
/// the call site — `F @x` is positional if `$x` is an array and by-name if it
/// is a hashtable — so both spellings reach the callee through this one helper.
/// Parameter names are matched case-insensitively, which
/// `splatting_case_insensitive_parameter_names` asserts (`@{ username = … }`
/// binding `$UserName`).
///
/// Stack: `[container, name, index]` → `[value]`.
pub fn emit_splat_arg(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let index = chunk.alloc_scratch(3);
    let name = index + 1;
    let container = name + 1;
    let is_array = chunk.add_import("ecma:array", "isArray");
    let cast_bool = chunk.add_import("wasm:js-boolean", "cast");
    let array_get = chunk.add_import("ecma:array", "get");

    chunk.emit_op_u16(Op::LOCAL_SET, index, line);
    chunk.emit_op_u16(Op::LOCAL_SET, name, line);
    chunk.emit_op_u16(Op::LOCAL_SET, container, line);

    chunk.emit_op_u16(Op::LOCAL_GET, container, line);
    chunk.emit_call(is_array, 1, line);
    chunk.emit_call(cast_bool, 1, line);
    chunk.emit_if_value(line);

    chunk.emit_op_u16(Op::LOCAL_GET, container, line);
    chunk.emit_op_u16(Op::LOCAL_GET, index, line);
    chunk.emit_call(array_get, 2, line);

    chunk.emit_else(line);

    chunk.emit_op_u16(Op::LOCAL_GET, container, line);
    chunk.emit_op_u16(Op::LOCAL_GET, name, line);
    vybe_compiler::primitives::strings::emit_to_lower(chunk, line);
    vybe_compiler::primitives::dict::emit_get_dynamic(chunks, current, line);

    chunks[current].emit_end(line);
}

/// PowerShell's `/`, which THROWS on a zero divisor rather than answering
/// infinity: `1 / 0` raises `RuntimeException` and is the idiomatic way a test
/// provokes an error for `trap` and `try`/`catch` to catch.
///
/// ⛔This is why the whole `exceptions_trap_statement_scope` family failed with
/// the trap never firing: `1 / 0` evaluated to `Infinity`, so nothing was
/// thrown and the function returned `Infinity AfterTrap` — two success-stream
/// values instead of one. The traps themselves were fine.
///
/// The exception is built with the SHARED `errors` primitives
/// (`emit_class_alloc` → `emit_exception_new_finalize` →
/// `emit_stamp_exception_ancestors` → `emit_throw`), so it carries the same
/// language-exception tag every `try_table` matches on and the ancestor stamp
/// that makes `catch [System.DivideByZeroException]` and its base types match.
///
/// ⚠ Known deviation: PowerShell decides by TYPE, so `[double]1 / 0` is
/// `Infinity` there and throws here. Both operands are `f64` by the time this
/// runs, so the distinction is not available — and no corpus test spells it.
///
/// Stack: `[a, b]` → `[quotient]`, or diverges.
pub fn emit_divide(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let b = chunk.alloc_scratch(2);
    let a = b + 1;
    chunk.emit_op_u16(Op::LOCAL_SET, b, line);
    chunk.emit_op_u16(Op::LOCAL_SET, a, line);

    let test_num = chunk.add_import("wasm:js-number", "test");
    let to_f64 = chunk.add_import("wasm:js-number", "toF64");
    chunk.emit_op_u16(Op::LOCAL_GET, b, line);
    chunk.emit_call(test_num, 1, line);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::LOCAL_GET, b, line);
    chunk.emit_call(to_f64, 1, line);
    chunk.emit_f64_const(0.0, line);
    chunk.emit_op(Op::F64_EQ, line);
    chunk.emit_if(line);

    let msg = chunk.alloc_scratch(1);
    chunk.emit_string_const("Attempted to divide by zero.", line);
    chunk.emit_op_u16(Op::LOCAL_SET, msg, line);
    // The shape `emit_exception_new_finalize` expects: [obj, obj, msg].
    vybe_compiler::primitives::class_slots::emit_class_alloc(chunk, line);
    chunk.emit_dup(line);
    chunk.emit_op_u16(Op::LOCAL_GET, msg, line);
    vybe_compiler::primitives::errors::emit_exception_new_finalize(
        chunk,
        "DivideByZeroException",
        line,
    );
    vybe_compiler::primitives::errors::emit_stamp_exception_ancestors(
        chunk,
        "DivideByZeroException",
        line,
    );
    vybe_compiler::primitives::errors::emit_throw(chunk, line);

    chunk.emit_end(line);
    chunk.emit_end(line);

    // ⛔NOT a bare `F64_DIV`. `/` still has to reach a USER operator method —
    // `[timespan]::FromHours(6) / [timespan]::FromHours(2)` is `3.0` through
    // `TimeSpan`'s own operator, and dividing the objects numerically answered
    // nonsense. `emit_rich_arithmetic` is the same helper the shared
    // `emit_rich_binop` uses, so the non-zero path behaves exactly as it did
    // before this adapter existed.
    vybe_compiler::primitives::expressions::emit_rich_arithmetic(
        chunk,
        a,
        b,
        &vybe_ast::protocol_slot_key(vybe_ast::ProtocolSlot::Div),
        |c: &mut Chunk, l: u32| c.emit_op(Op::F64_DIV, l),
        line,
    );
}

/// `.GetEnumerator()` — which iterates VALUES for an array and KEY/VALUE PAIRS
/// for a hashtable.
///
/// ⛔One `[value_methods]` row cannot say both, and it was saying the
/// hashtable answer for everything: `[int[]]@(1,2,3).GetEnumerator()` handed
/// back `0,1 1,2 2,3` — index/value pairs — so a class exposing
/// `[IEnumerator] GetEnumerator() { return $this.Items.GetEnumerator() }`
/// summed pairs and answered `NaN`. That is the whole
/// `classes_enumerable_and_iterator` family, and the class was never the
/// problem.
///
/// An array already IS its own enumeration here, so it is handed straight
/// back.
///
/// ⛔A HASHTABLE ENUMERATES `DictionaryEntry`, NOT PAIRS. `ecma:object:entries`
/// answers `[key, value]` ARRAYS, and `$e.Key` on an array is nothing — so
/// `$h.GetEnumerator() | … $_.Key` read empty for every entry. .NET's
/// `IDictionaryEnumerator` yields an entry carrying `Key` and `Value`, and the
/// corpus reads exactly those two names, so each pair is rebuilt as that
/// object.
///
/// Stack: `[receiver]` → `[enumerable]`.
pub fn emit_get_enumerator(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let recv = chunk.alloc_scratch(1);
    let pairs = chunk.alloc_scratch(1);
    let out = chunk.alloc_scratch(1);
    let cursor = chunk.alloc_scratch(1);
    let pair = chunk.alloc_scratch(1);
    let entry = chunk.alloc_scratch(1);
    let member = chunk.alloc_scratch(1);
    let is_array = chunk.add_import("ecma:array", "isArray");
    let cast_bool = chunk.add_import("wasm:js-boolean", "cast");
    let entries = chunk.add_import("ecma:object", "entries");
    let arr_len = chunk.add_import("ecma:array", "length");
    let arr_get = chunk.add_import("ecma:array", "get");
    let arr_push = chunk.add_import("ecma:array", "push");
    let obj_set = chunk.add_import("ecma:object", "set");
    let arr_new = chunk.add_import("ecma:array", "new");
    let obj_new = chunk.add_import("ecma:object", "new");

    chunk.emit_op_u16(Op::LOCAL_SET, recv, line);

    chunk.emit_op_u16(Op::LOCAL_GET, recv, line);
    chunk.emit_call(is_array, 1, line);
    chunk.emit_call(cast_bool, 1, line);
    chunk.emit_if_value(line);

    chunk.emit_op_u16(Op::LOCAL_GET, recv, line);

    chunk.emit_else(line);

    chunk.emit_op_u16(Op::LOCAL_GET, recv, line);
    chunk.emit_call(entries, 1, line);
    chunk.emit_op_u16(Op::LOCAL_SET, pairs, line);
    chunk.emit_call(arr_new, 0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, out, line);
    chunk.emit_i32_const(0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, cursor, line);

    chunk.emit_block(line);
    chunk.emit_loop_s(line);
    chunk.emit_op_u16(Op::LOCAL_GET, cursor, line);
    chunk.emit_op_u16(Op::LOCAL_GET, pairs, line);
    chunk.emit_call(arr_len, 1, line);
    chunk.emit_op(Op::I32_GE_S, line);
    chunk.emit_br_if(1, line);

    chunk.emit_op_u16(Op::LOCAL_GET, pairs, line);
    chunk.emit_op_u16(Op::LOCAL_GET, cursor, line);
    chunk.emit_call(arr_get, 2, line);
    chunk.emit_op_u16(Op::LOCAL_SET, pair, line);

    // ⛔THE ENTRY IS A CLASS SLOT OBJECT, NOT A PROPERTY BAG. A member read
    // compiled for a loop or script-block binding goes through the slot path,
    // and a key written with `ecma:object:set` is not there — `$_.Key` read
    // EMPTY inside a block while the identical `$a[0].Key` answered. This is
    // the shape `[pscustomobject]@{…}` already builds, which is why `$_.Code`
    // resolves on one of those.
    use vybe_compiler::primitives::class_slots::{
        ClassSlot, ObjSource, PlainNames, ValueSource, emit_class_alloc, emit_class_set, resolve,
    };
    emit_class_alloc(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_SET, entry, line);
    // ⛔THE KEY IS STORED FOLDED. PowerShell is case-insensitive, so the
    // compiler's own member resolution canonicalizes a field to lowercase
    // before it reads it; a slot written with the .NET spelling `Key` is not
    // the one `$_.Key` looks for, and the read answered EMPTY inside a loop or
    // script block while resolving outside one. `PlainNames` does not fold, so
    // the fold is applied here.
    for (index, field) in [(0, "key"), (1, "value")] {
        chunk.emit_op_u16(Op::LOCAL_GET, pair, line);
        chunk.emit_i32_const(index, line);
        chunk.emit_call(arr_get, 2, line);
        chunk.emit_op_u16(Op::LOCAL_SET, member, line);
        emit_class_set(
            chunk,
            ObjSource::Local(entry),
            &resolve(&ClassSlot::instance(field), &PlainNames),
            ValueSource::Local(member),
            line,
        );
    }
    chunk.emit_op_u16(Op::LOCAL_GET, out, line);
    chunk.emit_op_u16(Op::LOCAL_GET, entry, line);
    chunk.emit_call(arr_push, 2, line);
    chunk.emit_op(Op::DROP, line);

    chunk.emit_op_u16(Op::LOCAL_GET, cursor, line);
    chunk.emit_i32_const(1, line);
    chunk.emit_op(Op::I32_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_SET, cursor, line);
    chunk.emit_br(0, line);
    chunk.emit_end(line);
    chunk.emit_end(line);

    chunk.emit_op_u16(Op::LOCAL_GET, out, line);

    chunk.emit_end(line);
}

/// A pipeline's value: a stream of ONE hands back that one thing.
///
/// `@($obj) | ForEach-Object { $_ }` is the object, not a one-element array —
/// `$r.Tag` reaches the property in PowerShell and answered EMPTY here, because
/// the fold's result stayed wrapped. The same rule the success stream already
/// applies to a function's output, applied to a pipeline's.
///
/// ⛔Only for a length of exactly ONE. Zero stays an empty array and two or
/// more stay an array, which is what every `$x.Count` in the corpus expects.
///
/// Stack: `[value]` → `[value]`.
pub fn emit_unwrap_single(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let value = chunk.alloc_scratch(1);
    let is_array = chunk.add_import("ecma:array", "isArray");
    let cast_bool = chunk.add_import("wasm:js-boolean", "cast");
    let length = chunk.add_import("ecma:array", "length");
    let get = chunk.add_import("ecma:array", "get");

    chunk.emit_op_u16(Op::LOCAL_SET, value, line);

    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    chunk.emit_call(is_array, 1, line);
    chunk.emit_call(cast_bool, 1, line);
    chunk.emit_if_value(line);

    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    chunk.emit_call(length, 1, line);
    chunk.emit_i32_const(1, line);
    chunk.emit_op(Op::I32_EQ, line);
    chunk.emit_if_value(line);
    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    chunk.emit_i32_const(0, line);
    chunk.emit_call(get, 2, line);
    chunk.emit_else(line);
    // An EMPTY stream is `$null`, not an empty list. Measured against pwsh
    // 7.6.4: `$x = @(1,2) | Where-Object { $false }` leaves `$x` at `$null`,
    // while the literal `@()` — which never reaches this rule — stays an empty
    // array. Answering the empty array here made every "did this produce
    // nothing" test compare a list against `$null` and fail.
    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    chunk.emit_call(length, 1, line);
    chunk.emit_op(Op::I32_EQZ, line);
    chunk.emit_if_value(line);
    chunk.emit_ref_null(0x6e, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    chunk.emit_end(line);
    chunk.emit_end(line);

    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    chunk.emit_end(line);
}

/// `Get-Unique` — drop each item equal to the one before it.
///
/// ⛔ADJACENT, NOT DISTINCT. PowerShell documents `Get-Unique` as operating on a
/// SORTED list, so `1,2,1` keeps all three items and a set would silently
/// answer two. `-AsString` compares the rendered form instead of the value,
/// which is how `"1"` and `1` collapse.
///
/// Stack: `[items, as_string]` → `[array]`.
pub fn emit_unique_adjacent(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let on_type = chunk.alloc_scratch(7);
    let as_string = on_type + 1;
    let items = on_type + 2;
    let out = on_type + 3;
    let cursor = on_type + 4;
    let current_key = on_type + 5;
    let last_key = on_type + 6;

    let arr_new = chunk.add_import("ecma:array", "new");
    let arr_len = chunk.add_import("ecma:array", "length");
    let arr_get = chunk.add_import("ecma:array", "get");
    let arr_push = chunk.add_import("ecma:array", "push");
    let value_typeof = chunk.add_import("ecma:value", "typeof");

    chunk.emit_op_u16(Op::LOCAL_SET, on_type, line);
    chunk.emit_op_u16(Op::LOCAL_SET, as_string, line);
    chunk.emit_op_u16(Op::LOCAL_SET, items, line);
    chunk.emit_call(arr_new, 0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, out, line);
    chunk.emit_i32_const(0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, cursor, line);

    chunk.emit_block(line);
    chunk.emit_loop_s(line);
    chunk.emit_op_u16(Op::LOCAL_GET, cursor, line);
    chunk.emit_op_u16(Op::LOCAL_GET, items, line);
    chunk.emit_call(arr_len, 1, line);
    chunk.emit_op(Op::I32_GE_S, line);
    chunk.emit_br_if(1, line);

    // key = on_type ? typeof(item) : as_string ? String(item) : item
    chunk.emit_op_u16(Op::LOCAL_GET, on_type, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    chunk.emit_op_u16(Op::LOCAL_GET, items, line);
    chunk.emit_op_u16(Op::LOCAL_GET, cursor, line);
    chunk.emit_call(arr_get, 2, line);
    chunk.emit_call(value_typeof, 1, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, as_string, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    chunk.emit_op_u16(Op::LOCAL_GET, items, line);
    chunk.emit_op_u16(Op::LOCAL_GET, cursor, line);
    chunk.emit_call(arr_get, 2, line);
    let _ = chunk;
    super::display::emit_to_unique_key(chunks, current, line);
    let chunk = &mut chunks[current];
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, items, line);
    chunk.emit_op_u16(Op::LOCAL_GET, cursor, line);
    chunk.emit_call(arr_get, 2, line);
    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_op_u16(Op::LOCAL_SET, current_key, line);

    // The first item has no predecessor, so it always survives.
    chunk.emit_op_u16(Op::LOCAL_GET, cursor, line);
    chunk.emit_i32_const(0, line);
    chunk.emit_op(Op::I32_EQ, line);
    chunk.emit_if_value(line);
    chunk.emit_bool_const(true, line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, current_key, line);
    chunk.emit_op_u16(Op::LOCAL_GET, last_key, line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_not(chunk, line);
    chunk.emit_end(line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if(line);
    chunk.emit_op_u16(Op::LOCAL_GET, out, line);
    chunk.emit_op_u16(Op::LOCAL_GET, items, line);
    chunk.emit_op_u16(Op::LOCAL_GET, cursor, line);
    chunk.emit_call(arr_get, 2, line);
    chunk.emit_call(arr_push, 2, line);
    chunk.emit_op(Op::DROP, line);
    chunk.emit_end(line);

    chunk.emit_op_u16(Op::LOCAL_GET, current_key, line);
    chunk.emit_op_u16(Op::LOCAL_SET, last_key, line);
    chunk.emit_op_u16(Op::LOCAL_GET, cursor, line);
    chunk.emit_i32_const(1, line);
    chunk.emit_op(Op::I32_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_SET, cursor, line);
    chunk.emit_br(0, line);
    chunk.emit_end(line);
    chunk.emit_end(line);

    chunk.emit_op_u16(Op::LOCAL_GET, out, line);
}

/// `$obj.PSObject.Properties` / `.Members` — the object's ETS members, each one
/// a descriptor carrying `Name`, `Value` and `MemberType`.
///
/// The object model already answers the hard half: `Object.keys` folds an
/// accessor back to its bare name (`__get_c` is listed as `c`) and hides every
/// `__`-prefixed key, so the private `__ps_psobject` view and the raw accessor
/// slots never reach this list.
///
/// ⚠ THAT FILTER IS SHARED AND NOT DECLARED HERE. `platforms/ecma/src/object.rs`
/// drops `__`-prefixed own keys unconditionally, for every language — which is
/// an ECMA violation on the JS side (`Object.keys({__foo:1})` must list
/// `__foo`). If that is ever corrected, this list starts leaking
/// `__ps_psobject` and the raw `__get_`/`__set_` slots, and the filter has to
/// move HERE at the same time. What is left is naming each member's KIND,
/// which the accessor slots themselves record — a name with a `__get_` beside
/// it is a `ScriptProperty`, a name holding a function is a `ScriptMethod`, and
/// anything else is a `NoteProperty`.
///
/// ⛔A DESCRIPTOR IS A SNAPSHOT. `Add` and `Remove` write to the TARGET, not to
/// this list, and are recognized at their call sites for that reason.
///
/// Stack: `[target]` -> `[descriptors]`.
pub fn emit_psobject_properties(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    let target = chunk.alloc_scratch(10);
    let keys = target + 1;
    let out = target + 2;
    let cursor = target + 3;
    let key = target + 4;
    let getter = target + 5;
    let setter = target + 6;
    let desc = target + 7;
    let wanted = target + 8;
    let aliased = target + 9;

    let arr_new = chunk.add_import("ecma:array", "new");
    let arr_len = chunk.add_import("ecma:array", "length");
    let arr_get = chunk.add_import("ecma:array", "get");
    let arr_push = chunk.add_import("ecma:array", "push");
    let obj_keys = chunk.add_import("ecma:object", "keys");
    let obj_get = chunk.add_import("ecma:object", "get");
    let obj_set = chunk.add_import("ecma:object", "set");
    let type_of = chunk.add_import("ecma:value", "typeof");
    let lower = chunk.add_import("ecma:string", "toLowerCase");

    // `$obj.PSObject.Properties["Name"]` asks for ONE member. Storage folds a
    // member's spelling, so the name asked for is folded to match it.
    if argc == 2 {
        chunk.emit_call(lower, 1, line);
        chunk.emit_op_u16(Op::LOCAL_SET, wanted, line);
    }
    chunk.emit_op_u16(Op::LOCAL_SET, target, line);
    chunk.emit_op_u16(Op::LOCAL_GET, target, line);
    chunk.emit_call(obj_keys, 1, line);
    chunk.emit_op_u16(Op::LOCAL_SET, keys, line);
    chunk.emit_call(arr_new, 0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, out, line);
    chunk.emit_i32_const(0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, cursor, line);

    chunk.emit_block(line);
    chunk.emit_loop_s(line);
    chunk.emit_op_u16(Op::LOCAL_GET, cursor, line);
    chunk.emit_op_u16(Op::LOCAL_GET, keys, line);
    chunk.emit_call(arr_len, 1, line);
    chunk.emit_op(Op::I32_GE_S, line);
    chunk.emit_br_if(1, line);

    chunk.emit_op_u16(Op::LOCAL_GET, keys, line);
    chunk.emit_op_u16(Op::LOCAL_GET, cursor, line);
    chunk.emit_call(arr_get, 2, line);
    chunk.emit_op_u16(Op::LOCAL_SET, key, line);

    // The accessor slots beside the name, read RAW: `object.get` on `__get_X`
    // hands back the function, where a read of `X` would run it.
    chunk.emit_op_u16(Op::LOCAL_GET, target, line);
    chunk.emit_string_const("__get_", line);
    chunk.emit_op_u16(Op::LOCAL_GET, key, line);
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
    chunk.emit_call(obj_get, 2, line);
    chunk.emit_op_u16(Op::LOCAL_SET, getter, line);
    chunk.emit_op_u16(Op::LOCAL_GET, target, line);
    chunk.emit_string_const("__set_", line);
    chunk.emit_op_u16(Op::LOCAL_GET, key, line);
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
    chunk.emit_call(obj_get, 2, line);
    chunk.emit_op_u16(Op::LOCAL_SET, setter, line);
    chunk.emit_op_u16(Op::LOCAL_GET, target, line);
    chunk.emit_string_const("__alias_", line);
    chunk.emit_op_u16(Op::LOCAL_GET, key, line);
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
    chunk.emit_call(obj_get, 2, line);
    chunk.emit_op_u16(Op::LOCAL_SET, aliased, line);

    vybe_compiler::primitives::dict::emit_new(chunks, current, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, desc, line);

    // ⛔FOLDED SPELLINGS. PowerShell is case-insensitive, so a member read is
    // folded to lower case before it reaches storage; a field written here as
    // `Name` is stored under that spelling and `$p.Name` looks for `name` and
    // finds nothing. The descriptor is read far more often than it is printed.
    let field = |chunk: &mut Chunk, name: &str| {
        chunk.emit_op_u16(Op::LOCAL_GET, desc, line);
        chunk.emit_string_const(name, line);
    };

    let chunk = &mut chunks[current];

    field(chunk, "name");
    chunk.emit_op_u16(Op::LOCAL_GET, key, line);
    chunk.emit_call(obj_set, 3, line);
    chunk.emit_op(Op::DROP, line);

    // Reading the NAME runs a getter, which is what a descriptor's `Value` is.
    field(chunk, "value");
    chunk.emit_op_u16(Op::LOCAL_GET, target, line);
    chunk.emit_op_u16(Op::LOCAL_GET, key, line);
    chunk.emit_call(obj_get, 2, line);
    chunk.emit_call(obj_set, 3, line);
    chunk.emit_op(Op::DROP, line);

    // An alias forwards through a getter like a `ScriptProperty` does, so the
    // `__alias_` slot is what tells the two apart and is asked first.
    field(chunk, "membertype");
    chunk.emit_op_u16(Op::LOCAL_GET, aliased, line);
    chunk.emit_call(type_of, 1, line);
    chunk.emit_string_const("string", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    chunk.emit_string_const("AliasProperty", line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, getter, line);
    chunk.emit_call(type_of, 1, line);
    chunk.emit_string_const("function", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    chunk.emit_string_const("ScriptProperty", line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, target, line);
    chunk.emit_op_u16(Op::LOCAL_GET, key, line);
    chunk.emit_call(obj_get, 2, line);
    chunk.emit_call(type_of, 1, line);
    chunk.emit_string_const("function", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    chunk.emit_string_const("ScriptMethod", line);
    chunk.emit_else(line);
    chunk.emit_string_const("NoteProperty", line);
    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_call(obj_set, 3, line);
    chunk.emit_op(Op::DROP, line);

    // A plain value reads and writes; a script member does whichever half it
    // was given a block for.
    field(chunk, "isgettable");
    chunk.emit_op_u16(Op::LOCAL_GET, setter, line);
    chunk.emit_call(type_of, 1, line);
    chunk.emit_string_const("function", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    chunk.emit_op_u16(Op::LOCAL_GET, getter, line);
    chunk.emit_call(type_of, 1, line);
    chunk.emit_string_const("function", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    chunk.emit_else(line);
    chunk.emit_bool_const(true, line);
    chunk.emit_end(line);
    chunk.emit_call(obj_set, 3, line);
    chunk.emit_op(Op::DROP, line);

    field(chunk, "issettable");
    chunk.emit_op_u16(Op::LOCAL_GET, getter, line);
    chunk.emit_call(type_of, 1, line);
    chunk.emit_string_const("function", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    chunk.emit_op_u16(Op::LOCAL_GET, setter, line);
    chunk.emit_call(type_of, 1, line);
    chunk.emit_string_const("function", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    chunk.emit_else(line);
    chunk.emit_bool_const(true, line);
    chunk.emit_end(line);
    chunk.emit_call(obj_set, 3, line);
    chunk.emit_op(Op::DROP, line);

    field(chunk, "getterscript");
    chunk.emit_op_u16(Op::LOCAL_GET, getter, line);
    chunk.emit_call(obj_set, 3, line);
    chunk.emit_op(Op::DROP, line);

    field(chunk, "setterscript");
    chunk.emit_op_u16(Op::LOCAL_GET, setter, line);
    chunk.emit_call(obj_set, 3, line);
    chunk.emit_op(Op::DROP, line);

    if argc == 2 {
        chunk.emit_op_u16(Op::LOCAL_GET, key, line);
        chunk.emit_op_u16(Op::LOCAL_GET, wanted, line);
        vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
        vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
        chunk.emit_if(line);
    }
    chunk.emit_op_u16(Op::LOCAL_GET, out, line);
    chunk.emit_op_u16(Op::LOCAL_GET, desc, line);
    chunk.emit_call(arr_push, 2, line);
    chunk.emit_op(Op::DROP, line);
    if argc == 2 {
        chunk.emit_end(line);
    }

    chunk.emit_op_u16(Op::LOCAL_GET, cursor, line);
    chunk.emit_i32_const(1, line);
    chunk.emit_op(Op::I32_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_SET, cursor, line);
    chunk.emit_br(0, line);
    chunk.emit_end(line);
    chunk.emit_end(line);

    // One member was asked for: that member, or null when the object carries no
    // such member — never a one-element list.
    if argc == 2 {
        chunk.emit_op_u16(Op::LOCAL_GET, out, line);
        chunk.emit_call(arr_len, 1, line);
        chunk.emit_i32_const(0, line);
        chunk.emit_op(Op::I32_GT_S, line);
        chunk.emit_if_value(line);
        chunk.emit_op_u16(Op::LOCAL_GET, out, line);
        chunk.emit_i32_const(0, line);
        chunk.emit_call(arr_get, 2, line);
        chunk.emit_else(line);
        chunk.emit_ref_null(0x6e, line);
        chunk.emit_end(line);
        return;
    }

    chunk.emit_op_u16(Op::LOCAL_GET, out, line);
}

/// `-eq` / `-ne` — PowerShell's comparison, which folds case.
///
/// ⛔A COMPARISON IS CASE-INSENSITIVE UNLESS THE OPERATOR SAYS OTHERWISE.
/// Measured against pwsh 7.6.4: `'abc' -eq 'ABC'` is True and `-ceq` is the
/// case-sensitive spelling. The shared `emit_dyn_eq` compares exactly, so every
/// unprefixed string comparison in the language answered False for operands
/// PowerShell calls equal.
///
/// Only a comparison of TWO STRINGS folds. Everything else — numbers, objects,
/// `$null`, a string against a number — keeps the shared equality, which is
/// what makes `2 -eq 2.0` True and keeps the numeric coercion intact.
///
/// Stack: `[a, b]` → `[bool]`.
pub fn emit_case_folding_eq(chunks: &mut [Chunk], current: usize, negated: bool, line: u32) {
    let chunk = &mut chunks[current];
    let b_slot = chunk.alloc_scratch(2);
    let a_slot = b_slot + 1;

    let is_string = chunk.add_import("wasm:js-string", "test");
    let lower = chunk.add_import("ecma:string", "toLowerCase");

    chunk.emit_op_u16(Op::LOCAL_SET, b_slot, line);
    chunk.emit_op_u16(Op::LOCAL_SET, a_slot, line);

    chunk.emit_op_u16(Op::LOCAL_GET, a_slot, line);
    chunk.emit_call(is_string, 1, line);
    chunk.emit_op_u16(Op::LOCAL_GET, b_slot, line);
    chunk.emit_call(is_string, 1, line);
    chunk.emit_op(Op::I32_AND, line);
    chunk.emit_if_value(line);

    chunk.emit_op_u16(Op::LOCAL_GET, a_slot, line);
    chunk.emit_call(lower, 1, line);
    chunk.emit_op_u16(Op::LOCAL_GET, b_slot, line);
    chunk.emit_call(lower, 1, line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);

    chunk.emit_else(line);

    // ⛔NOT `emit_dyn_eq`. PowerShell coerces the right operand to the LEFT
    // one's type, so `2 -eq '2'` is True — the same left-operand rule `+` has.
    // The exact comparison answered False, and `Get-Date -UFormat %j` compared
    // against `"130"` was one of six tests that lost that coercion.
    chunk.emit_op_u16(Op::LOCAL_GET, a_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, b_slot, line);
    let abstract_eq = chunk.add_import("ecma:value", "abstractEq");
    chunk.emit_call(abstract_eq, 2, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);

    chunk.emit_end(line);

    if negated {
        chunk.emit_op(Op::I32_EQZ, line);
    }

    // A real boolean, not the i32 the comparison leaves: under
    // `materialize_bool_results` an i32 `1` does not compare equal to the
    // language's `$true`, and it renders as `1` rather than `True`.
    chunk.emit_if_value(line);
    chunk.emit_bool_const(true, line);
    chunk.emit_else(line);
    chunk.emit_bool_const(false, line);
    chunk.emit_end(line);
}

/// `Compare-Object -ReferenceObject $a -DifferenceObject $b` — a MULTISET
/// difference, tagged with which side each item came from.
///
/// PowerShell compares by VALUE, not by identity, and counts duplicates: two
/// `"dup"` on the left against one on the right leaves one `<=`. So each item is
/// reduced to a comparison KEY — its display text, or its `-Property` values
/// joined — and matching consumes a slot on each side rather than testing
/// membership.
///
/// Case folds unless `-CaseSensitive`, the same rule every other PowerShell
/// comparison follows.
///
/// Stack: `[ref, diff, include_equal, exclude_different, case_sensitive,
/// properties, pass_thru]` → `[entries]`.
pub fn emit_compare_object(chunks: &mut [Chunk], current: usize, line: u32) {
    let pass_thru = chunks[current].alloc_scratch(16);
    let props = pass_thru + 1;
    let case_sensitive = pass_thru + 2;
    let exclude_different = pass_thru + 3;
    let include_equal = pass_thru + 4;
    let diff = pass_thru + 5;
    let reference = pass_thru + 6;
    let ref_keys = pass_thru + 7;
    let diff_keys = pass_thru + 8;
    let ref_used = pass_thru + 9;
    let diff_used = pass_thru + 10;
    let out = pass_thru + 11;
    let i = pass_thru + 12;
    let j = pass_thru + 13;
    let key = pass_thru + 14;
    let matched = pass_thru + 15;

    let arr_new = chunks[current].add_import("ecma:array", "new");
    let arr_len = chunks[current].add_import("ecma:array", "length");
    let arr_get = chunks[current].add_import("ecma:array", "get");
    let arr_set = chunks[current].add_import("ecma:array", "set");
    let arr_push = chunks[current].add_import("ecma:array", "push");
    let obj_get = chunks[current].add_import("ecma:object", "get");
    let obj_set = chunks[current].add_import("ecma:object", "set");
    let lower = chunks[current].add_import("ecma:string", "toLowerCase");

    for slot in [
        pass_thru,
        props,
        case_sensitive,
        exclude_different,
        include_equal,
        diff,
        reference,
    ] {
        chunks[current].emit_op_u16(Op::LOCAL_SET, slot, line);
    }

    // One item's comparison key, left on the stack.
    let emit_key = |chunks: &mut [Chunk], current: usize, item: u16| {
        let p = chunks[current].alloc_scratch(2);
        let acc = p + 1;
        // `-Property Name, Age` compares those members rather than the whole
        // object; a NUL between them keeps `("a","bc")` and `("ab","c")` apart.
        chunks[current].emit_op_u16(Op::LOCAL_GET, props, line);
        chunks[current].emit_op(Op::REF_IS_NULL, line);
        chunks[current].emit_if_value(line);

        chunks[current].emit_op_u16(Op::LOCAL_GET, item, line);
        super::display::emit_to_display(chunks, current, line);

        chunks[current].emit_else(line);

        chunks[current].emit_string_const("", line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, acc, line);
        chunks[current].emit_i32_const(0, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, p, line);
        chunks[current].emit_block(line);
        chunks[current].emit_loop_s(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, p, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, props, line);
        chunks[current].emit_call(arr_len, 1, line);
        chunks[current].emit_op(Op::I32_GE_S, line);
        chunks[current].emit_br_if(1, line);

        chunks[current].emit_op_u16(Op::LOCAL_GET, acc, line);
        chunks[current].emit_string_const("\u{0}", line);
        vybe_compiler::primitives::ops::emit_dyn_add(&mut chunks[current], line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, item, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, props, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, p, line);
        chunks[current].emit_call(arr_get, 2, line);
        // ⛔THE MEMBER NAME IS FOLDED IN STORAGE. `-Property Org` reads the key
        // `org`, so asking for `Org` missed and every object produced the same
        // empty composite key — which made every pair compare EQUAL.
        chunks[current].emit_call(lower, 1, line);
        chunks[current].emit_call(obj_get, 2, line);
        super::display::emit_to_display(chunks, current, line);
        vybe_compiler::primitives::ops::emit_dyn_add(&mut chunks[current], line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, acc, line);

        chunks[current].emit_op_u16(Op::LOCAL_GET, p, line);
        chunks[current].emit_i32_const(1, line);
        chunks[current].emit_op(Op::I32_ADD, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, p, line);
        chunks[current].emit_br(0, line);
        chunks[current].emit_end(line);
        chunks[current].emit_end(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, acc, line);

        chunks[current].emit_end(line);

        // The fold, unless the caller asked for an exact comparison.
        chunks[current].emit_op_u16(Op::LOCAL_SET, acc, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, case_sensitive, line);
        vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
        chunks[current].emit_if_value(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, acc, line);
        chunks[current].emit_else(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, acc, line);
        chunks[current].emit_call(lower, 1, line);
        chunks[current].emit_end(line);
    };

    // `keys(source) -> [key…]`, and a parallel `used` array of falses.
    let mut build = |chunks: &mut [Chunk], current: usize, source: u16, keys: u16, used: u16| {
        chunks[current].emit_call(arr_new, 0, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, keys, line);
        chunks[current].emit_call(arr_new, 0, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, used, line);
        chunks[current].emit_i32_const(0, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
        chunks[current].emit_block(line);
        chunks[current].emit_loop_s(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, source, line);
        chunks[current].emit_call(arr_len, 1, line);
        chunks[current].emit_op(Op::I32_GE_S, line);
        chunks[current].emit_br_if(1, line);

        chunks[current].emit_op_u16(Op::LOCAL_GET, source, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
        chunks[current].emit_call(arr_get, 2, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, key, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, keys, line);
        emit_key(chunks, current, key);
        chunks[current].emit_call(arr_push, 2, line);
        chunks[current].emit_op(Op::DROP, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, used, line);
        chunks[current].emit_bool_const(false, line);
        chunks[current].emit_call(arr_push, 2, line);
        chunks[current].emit_op(Op::DROP, line);

        chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
        chunks[current].emit_i32_const(1, line);
        chunks[current].emit_op(Op::I32_ADD, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
        chunks[current].emit_br(0, line);
        chunks[current].emit_end(line);
        chunks[current].emit_end(line);
    };
    build(chunks, current, reference, ref_keys, ref_used);
    build(chunks, current, diff, diff_keys, diff_used);

    chunks[current].emit_call(arr_new, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);

    // One output entry. `-PassThru` hands back the ORIGINAL object with the
    // indicator attached, which is what lets the caller keep its own members.
    let emit_entry = |chunks: &mut [Chunk], current: usize, item: u16, indicator: &str| {
        let entry = chunks[current].alloc_scratch(1);
        chunks[current].emit_op_u16(Op::LOCAL_GET, pass_thru, line);
        vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
        chunks[current].emit_if_value(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, item, line);
        chunks[current].emit_else(line);
        vybe_compiler::primitives::dict::emit_new(chunks, current, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, entry, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, entry, line);
        // ⛔FOLDED KEYS. A member read is folded before it reaches storage, so
        // `InputObject` written here is not what `$_.InputObject` looks for.
        chunks[current].emit_string_const("inputobject", line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, item, line);
        chunks[current].emit_call(obj_set, 3, line);
        chunks[current].emit_op(Op::DROP, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, entry, line);
        chunks[current].emit_end(line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, entry, line);

        // `-Property Id` puts the COMPARED members on the output record —
        // that is what the caller reads them back off. Without them the record
        // carried only `InputObject`, so `$diff.UserId` answered empty.
        let q = chunks[current].alloc_scratch(1);
        chunks[current].emit_op_u16(Op::LOCAL_GET, props, line);
        chunks[current].emit_op(Op::REF_IS_NULL, line);
        chunks[current].emit_op(Op::I32_EQZ, line);
        chunks[current].emit_if(line);
        chunks[current].emit_i32_const(0, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, q, line);
        chunks[current].emit_block(line);
        chunks[current].emit_loop_s(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, q, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, props, line);
        chunks[current].emit_call(arr_len, 1, line);
        chunks[current].emit_op(Op::I32_GE_S, line);
        chunks[current].emit_br_if(1, line);

        chunks[current].emit_op_u16(Op::LOCAL_GET, entry, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, props, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, q, line);
        chunks[current].emit_call(arr_get, 2, line);
        chunks[current].emit_call(lower, 1, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, item, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, props, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, q, line);
        chunks[current].emit_call(arr_get, 2, line);
        chunks[current].emit_call(lower, 1, line);
        chunks[current].emit_call(obj_get, 2, line);
        chunks[current].emit_call(obj_set, 3, line);
        chunks[current].emit_op(Op::DROP, line);

        chunks[current].emit_op_u16(Op::LOCAL_GET, q, line);
        chunks[current].emit_i32_const(1, line);
        chunks[current].emit_op(Op::I32_ADD, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, q, line);
        chunks[current].emit_br(0, line);
        chunks[current].emit_end(line);
        chunks[current].emit_end(line);
        chunks[current].emit_end(line);

        chunks[current].emit_op_u16(Op::LOCAL_GET, entry, line);
        chunks[current].emit_string_const("sideindicator", line);
        chunks[current].emit_string_const(indicator, line);
        chunks[current].emit_call(obj_set, 3, line);
        chunks[current].emit_op(Op::DROP, line);

        chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, entry, line);
        chunks[current].emit_call(arr_push, 2, line);
        chunks[current].emit_op(Op::DROP, line);
    };

    // Pass one: pair each difference item with an unconsumed reference item.
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, j, line);
    chunks[current].emit_block(line);
    chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, j, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, diff, line);
    chunks[current].emit_call(arr_len, 1, line);
    chunks[current].emit_op(Op::I32_GE_S, line);
    chunks[current].emit_br_if(1, line);

    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, matched, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    chunks[current].emit_block(line);
    chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, reference, line);
    chunks[current].emit_call(arr_len, 1, line);
    chunks[current].emit_op(Op::I32_GE_S, line);
    chunks[current].emit_br_if(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, matched, line);
    chunks[current].emit_br_if(1, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, ref_used, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_call(arr_get, 2, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, ref_keys, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_call(arr_get, 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, diff_keys, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, j, line);
    chunks[current].emit_call(arr_get, 2, line);
    vybe_compiler::primitives::ops::emit_dyn_eq(&mut chunks[current], line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_op(Op::I32_AND, line);
    chunks[current].emit_if(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, ref_used, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_bool_const(true, line);
    // `array.set` pushes a value even though its signature declares no result
    // — `arrays.rs` drops after it for the same reason.
    chunks[current].emit_call(arr_set, 3, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, diff_used, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, j, line);
    chunks[current].emit_bool_const(true, line);
    // `array.set` pushes a value even though its signature declares no result
    // — `arrays.rs` drops after it for the same reason.
    chunks[current].emit_call(arr_set, 3, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, matched, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, include_equal, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
    chunks[current].emit_if(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, diff, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, j, line);
    chunks[current].emit_call(arr_get, 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, key, line);
    emit_entry(chunks, current, key, "==");
    chunks[current].emit_end(line);

    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, j, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, j, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);

    // Pass two: whatever each side did not pair off. `-ExcludeDifferent` keeps
    // only the matches, so both loops are gated on it.
    let mut unmatched =
        |chunks: &mut [Chunk], current: usize, source: u16, used: u16, indicator: &str| {
            chunks[current].emit_op_u16(Op::LOCAL_GET, exclude_different, line);
            vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
            chunks[current].emit_op(Op::I32_EQZ, line);
            chunks[current].emit_if(line);
            chunks[current].emit_i32_const(0, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
            chunks[current].emit_block(line);
            chunks[current].emit_loop_s(line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, source, line);
            chunks[current].emit_call(arr_len, 1, line);
            chunks[current].emit_op(Op::I32_GE_S, line);
            chunks[current].emit_br_if(1, line);

            chunks[current].emit_op_u16(Op::LOCAL_GET, used, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
            chunks[current].emit_call(arr_get, 2, line);
            vybe_compiler::primitives::ops::emit_dyn_to_bool(&mut chunks[current], line);
            chunks[current].emit_op(Op::I32_EQZ, line);
            chunks[current].emit_if(line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, source, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
            chunks[current].emit_call(arr_get, 2, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, key, line);
            emit_entry(chunks, current, key, indicator);
            chunks[current].emit_end(line);

            chunks[current].emit_op_u16(Op::LOCAL_GET, i, line);
            chunks[current].emit_i32_const(1, line);
            chunks[current].emit_op(Op::I32_ADD, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, i, line);
            chunks[current].emit_br(0, line);
            chunks[current].emit_end(line);
            chunks[current].emit_end(line);
            chunks[current].emit_end(line);
        };
    unmatched(chunks, current, diff, diff_used, "=>");
    unmatched(chunks, current, reference, ref_used, "<=");

    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
}

/// `$x.GetType()` — the value's type, whichever model the value came from.
///
/// Two emitters each answer half. `ref_typeof` is the class model's own: a
/// user-class instance answers its class as a Type with `.Name`. The dotnet
/// `get_type` answers the .NET VALUE types — a string is `String`, an int is
/// `Int32` — and is `undefined is not callable` on a class instance. Bound to
/// either one alone, half the corpus lost `GetType()`: measured, the dotnet
/// binding trapped on every `[C]::new().GetType()`, and the class binding
/// answered `""` for `"str".GetType().Name`.
///
/// A class instance is the only value `typeof` calls `object` here that is not
/// also a .NET collection, so that is the fork.
///
/// Stack: `[value]` → `[type]`.
pub fn emit_get_type(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let value = chunks[current].alloc_scratch(1);
    chunks[current].emit_op_u16(Op::LOCAL_SET, value, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, value, line);
    vybe_platform_dotnet::emitter::dispatch::dispatch("dotnet.get_type", chunks, current, 1, line);
}

fn emit_ps_type_object_from_name(chunk: &mut Chunk, full_name: u16, line: u32) {
    let split = chunk.add_import("ecma:string", "split");
    let at = chunk.add_import("ecma:array", "at");
    let obj = chunk.alloc_scratch(1);
    let short = chunk.alloc_scratch(1);

    vybe_compiler::primitives::class_slots::emit_class_alloc(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_SET, obj, line);

    chunk.emit_op_u16(Op::LOCAL_GET, full_name, line);
    chunk.emit_string_const(".", line);
    chunk.emit_call(split, 2, line);
    chunk.emit_i32_const(-1, line);
    chunk.emit_call(at, 2, line);
    chunk.emit_op_u16(Op::LOCAL_SET, short, line);

    set_ps_type_object_slot(chunk, obj, "Name", short, line);
    set_ps_type_object_slot(chunk, obj, "name", short, line);
    set_ps_type_object_slot(chunk, obj, "FullName", full_name, line);
    set_ps_type_object_slot(chunk, obj, "fullname", full_name, line);
    chunk.emit_op_u16(Op::LOCAL_GET, obj, line);
}

fn set_ps_type_object_slot(chunk: &mut Chunk, obj: u16, key: &str, value: u16, line: u32) {
    vybe_compiler::primitives::class_slots::emit_class_set(
        chunk,
        vybe_compiler::primitives::class_slots::ObjSource::Local(obj),
        &vybe_compiler::primitives::class_slots::resolve(
            &vybe_compiler::primitives::class_slots::ClassSlot::internal(key),
            &vybe_compiler::primitives::class_slots::PlainNames,
        ),
        vybe_compiler::primitives::class_slots::ValueSource::Local(value),
        line,
    );
}

/// `.Remove(…)` on an UNTYPED receiver — three receivers share the spelling at
/// one arity, and only the value at run time can tell them apart:
///
/// | receiver | `Remove(x)` |
/// |---|---|
/// | hashtable (`@{…}`) | delete the KEY |
/// | array / ArrayList | remove the first element equal to `x` |
/// | string | `'abc'.Remove(1)` is the text BEFORE index 1 |
///
/// The same fallback rule as `emit_collection_add`: a receiver whose type is
/// known at compile time — a `List[T]`, a `Hashtable` built by `::new()` —
/// resolves through the tree first and never reaches this row, so the tree's
/// leaves are not shadowed. A spelling-only rewrite was rejected here before
/// for exactly the reason this tests the shape instead.
///
/// Stack: `[recv, arg]` → `[result]`.
pub fn emit_collection_remove(chunks: &mut [Chunk], current: usize, line: u32) {
    let arg = chunks[current].alloc_scratch(2);
    let recv = arg + 1;
    let is_array = chunks[current].add_import("ecma:array", "isArray");
    let cast_bool = chunks[current].add_import("wasm:js-boolean", "cast");
    let is_string = chunks[current].add_import("wasm:js-string", "test");
    let obj_delete = chunks[current].add_import("ecma:object", "delete");

    chunks[current].emit_op_u16(Op::LOCAL_SET, arg, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, recv, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    chunks[current].emit_call(is_array, 1, line);
    chunks[current].emit_call(cast_bool, 1, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arg, line);
    vybe_compiler::primitives::collections::emit_remove_value(chunks, current, line);
    chunks[current].emit_else(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    chunks[current].emit_call(is_string, 1, line);
    chunks[current].emit_if_value(line);
    // `substring(0, index)` — the text before the cut.
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arg, line);
    vybe_compiler::primitives::strings::emit_substring(&mut chunks[current], line);
    chunks[current].emit_else(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, arg, line);
    chunks[current].emit_call(obj_delete, 2, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);
}

/// `.Clear()` on an UNTYPED receiver — the public `[value_methods]` row, reached
/// by a collection passed through a parameter, a field or a pipeline (a typed
/// receiver resolves through the tree first; a bare `@{…}` is renamed to
/// `__ps_ht_clear` by the walker and never gets here).
///
/// A `List`, `Stack`, `Queue` or `ArrayList` is array-kind at run time, so it
/// is emptied in place; anything else is a keyed collection whose own keys are
/// removed. `collections.clear_keyed` already answers null for an array and
/// would leave it FULL, which is what cleared nothing on a `HashSet`.
///
/// Stack: `[recv]` → `[null]`.
pub fn emit_collection_clear(chunks: &mut [Chunk], current: usize, line: u32) {
    let recv = chunks[current].alloc_scratch(1);
    let is_array = chunks[current].add_import("ecma:array", "isArray");
    let cast_bool = chunks[current].add_import("wasm:js-boolean", "cast");

    chunks[current].emit_op_u16(Op::LOCAL_SET, recv, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    chunks[current].emit_call(is_array, 1, line);
    chunks[current].emit_call(cast_bool, 1, line);
    chunks[current].emit_if_value(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    vybe_compiler::primitives::collections::emit_clear(chunks, current, line);
    chunks[current].emit_else(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, recv, line);
    vybe_compiler::primitives::collections::emit_clear_keyed(chunks, current, line);
    chunks[current].emit_end(line);
}

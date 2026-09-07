//! PowerShell's ONE value → string coercion.
//!
//! Bound as `[builtin_slots.string] to_string`, the slot the shared
//! interpolation arm reads (`expressions.rs`) — the same slot PHP declares to
//! render `null` as `""`. Interpolation used to fall through to
//! `strings::emit_to_string`, which is the ECMA coercion, so `"$null"` printed
//! `null` and `"$arr"` printed a comma-joined list.
//!
//! Verified against real `pwsh`:
//!
//! | value | PowerShell |
//! |---|---|
//! | `$null` | `""` |
//! | `$true` / `$false` | `True` / `False` — capitalized, unlike ECMA |
//! | `@(1,2,3)` | `1 2 3` — joined by `$OFS`, which defaults to a SPACE |

use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

/// Stack: `[value]` → `[string]`.
pub fn emit_to_display(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let v = chunk.alloc_scratch(1);

    let test_bool = chunk.add_import("wasm:js-boolean", "test");
    let is_array = chunk.add_import("ecma:array", "isArray");
    let cast_bool = chunk.add_import("wasm:js-boolean", "cast");
    let join = chunk.add_import("ecma:array", "join");

    chunk.emit_op_u16(Op::LOCAL_SET, v, line);

    // `$null` renders as nothing at all, not as the text "null".
    chunk.emit_op_u16(Op::LOCAL_GET, v, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if_value(line);
    chunk.emit_string_const("", line);
    chunk.emit_else(line);

    // PowerShell capitalizes its booleans; the ECMA coercion does not.
    chunk.emit_op_u16(Op::LOCAL_GET, v, line);
    chunk.emit_call(test_bool, 1, line);
    chunk.emit_if_value(line);
    chunk.emit_op_u16(Op::LOCAL_GET, v, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);
    chunk.emit_string_const("True", line);
    chunk.emit_else(line);
    chunk.emit_string_const("False", line);
    chunk.emit_end(line);

    chunk.emit_else(line);

    // `"$arr"` joins with `$OFS`. The default is a space — the ECMA coercion's
    // comma is what `"$(1,2,3)"` printed before this binding existed.
    chunk.emit_op_u16(Op::LOCAL_GET, v, line);
    chunk.emit_call(is_array, 1, line);
    chunk.emit_call(cast_bool, 1, line);
    chunk.emit_if_value(line);
    chunk.emit_op_u16(Op::LOCAL_GET, v, line);
    chunk.emit_string_const(" ", line);
    chunk.emit_call(join, 2, line);

    chunk.emit_else(line);

    // Everything else keeps the .NET rendering this slot was already bound to
    // (`common:dotnet.tostring_runtime`) — number formatting included. Only the
    // three cases above are PowerShell's own, so only they are handled here.
    chunk.emit_op_u16(Op::LOCAL_GET, v, line);
    vybe_platform_dotnet::emitter::core::runtime_adapter::emit_helper(
        "dotnet.tostring_runtime",
        std::slice::from_mut(chunk),
        0,
        1,
        line,
    );

    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_end(line);
}

/// `Get-Unique -AsString` key coercion.
///
/// PowerShell compares formatted object content here, not JavaScript's
/// `[object Object]`. For class-slot-backed PSCustomObjects, use the shared
/// dynamic slot owner to read each visible property and build a stable
/// `name=value;...` key. Non-objects keep the normal display coercion.
///
/// Stack: `[value]` -> `[string]`.
pub fn emit_to_unique_key(chunks: &mut [Chunk], current: usize, line: u32) {
    let chunk = &mut chunks[current];
    let v = chunk.alloc_scratch(6);
    let keys = v + 1;
    let out = v + 2;
    let cursor = v + 3;
    let key = v + 4;
    let value = v + 5;

    let type_of = chunk.add_import("ecma:value", "typeof");
    let is_array = chunk.add_import("ecma:array", "isArray");
    let cast_bool = chunk.add_import("wasm:js-boolean", "cast");
    let obj_keys = chunk.add_import("ecma:object", "keys");
    let arr_len = chunk.add_import("ecma:array", "length");
    let arr_get = chunk.add_import("ecma:array", "get");

    chunk.emit_op_u16(Op::LOCAL_SET, v, line);

    chunk.emit_op_u16(Op::LOCAL_GET, v, line);
    chunk.emit_op(Op::REF_IS_NULL, line);
    chunk.emit_if_value(line);
    chunk.emit_op_u16(Op::LOCAL_GET, v, line);
    let _ = chunk;
    emit_to_display(chunks, current, line);
    let chunk = &mut chunks[current];
    chunk.emit_else(line);

    chunk.emit_op_u16(Op::LOCAL_GET, v, line);
    chunk.emit_call(type_of, 1, line);
    chunk.emit_string_const("object", line);
    vybe_compiler::primitives::ops::emit_dyn_eq(chunk, line);
    vybe_compiler::primitives::ops::emit_dyn_to_bool(chunk, line);
    chunk.emit_if_value(line);

    chunk.emit_op_u16(Op::LOCAL_GET, v, line);
    chunk.emit_call(is_array, 1, line);
    chunk.emit_call(cast_bool, 1, line);
    chunk.emit_if_value(line);
    chunk.emit_op_u16(Op::LOCAL_GET, v, line);
    let _ = chunk;
    emit_to_display(chunks, current, line);
    let chunk = &mut chunks[current];
    chunk.emit_else(line);

    chunk.emit_op_u16(Op::LOCAL_GET, v, line);
    chunk.emit_call(obj_keys, 1, line);
    chunk.emit_op_u16(Op::LOCAL_SET, keys, line);
    chunk.emit_string_const("", line);
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

    vybe_compiler::primitives::class_slots::emit_class_get(
        chunk,
        vybe_compiler::primitives::class_slots::ObjSource::Local(v),
        &vybe_compiler::primitives::class_slots::resolve(
            &vybe_compiler::primitives::class_slots::ClassSlot::Dynamic(
                vybe_compiler::primitives::class_slots::ValueSource::Local(key),
            ),
            &vybe_compiler::primitives::class_slots::PlainNames,
        ),
        vybe_compiler::primitives::class_slots::Dest::Local(value),
        line,
    );

    chunk.emit_op_u16(Op::LOCAL_GET, out, line);
    chunk.emit_op_u16(Op::LOCAL_GET, key, line);
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
    chunk.emit_string_const("=", line);
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value, line);
    let _ = chunk;
    emit_to_display(chunks, current, line);
    let chunk = &mut chunks[current];
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
    chunk.emit_string_const(";", line);
    vybe_compiler::primitives::ops::emit_dyn_add(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_SET, out, line);

    chunk.emit_op_u16(Op::LOCAL_GET, cursor, line);
    chunk.emit_i32_const(1, line);
    chunk.emit_op(Op::I32_ADD, line);
    chunk.emit_op_u16(Op::LOCAL_SET, cursor, line);
    chunk.emit_br(0, line);
    chunk.emit_end(line);
    chunk.emit_end(line);
    chunk.emit_op_u16(Op::LOCAL_GET, out, line);

    chunk.emit_end(line);
    chunk.emit_else(line);
    chunk.emit_op_u16(Op::LOCAL_GET, v, line);
    let _ = chunk;
    emit_to_display(chunks, current, line);
    let chunk = &mut chunks[current];
    chunk.emit_end(line);
    chunk.emit_end(line);
}

/// Which shape `Format-List` / `Format-Table` / `Format-Wide` render.
#[derive(Clone, Copy, PartialEq, Eq)]
pub enum FormatMode {
    /// One `Name : Value` line per property, a blank line between objects.
    List,
    /// A header naming the properties, an underline, then a row per object.
    Table,
    /// A table row shape without the header and underline rows.
    TableNoHeaders,
    /// The objects' values alone, one per line.
    Wide,
}

/// The `Format-*` cmdlets — one renderer, three shapes.
///
/// A formatter is a SELECTOR over machinery that already exists: `Object.keys`
/// names an object's members (folding accessors back to their bare names and
/// hiding every `__` key), and `emit_to_display` is PowerShell's one value →
/// string coercion, so `$true` renders `True` and `$null` renders empty here
/// exactly as it does in an interpolation.
///
/// ⛔COLUMNS ARE NOT ALIGNED TO A MEASURED WIDTH. pwsh pads each column to its
/// widest cell and truncates past the console width with an ellipsis; the
/// corpus asserts CONTENT — `$output -match "Name\s*:\s*Value"` — so the
/// separator is a run of spaces that satisfies `\s*` rather than a layout
/// engine. The two tests that assert truncation fail rather than being told a
/// width this compiler has no console to ask for.
///
/// Stack: `[items]` → `[text]`.
pub fn emit_format(chunks: &mut [Chunk], current: usize, mode: FormatMode, line: u32) {
    let items = chunks[current].alloc_scratch(8);
    let out = items + 1;
    let cursor = items + 2;
    let keys = items + 3;
    let key_cursor = items + 4;
    let item = items + 5;
    let key = items + 6;
    let wrapped = items + 7;

    let arr_new = chunks[current].add_import("ecma:array", "new");
    let arr_len = chunks[current].add_import("ecma:array", "length");
    let arr_get = chunks[current].add_import("ecma:array", "get");
    let arr_push = chunks[current].add_import("ecma:array", "push");
    let is_array = chunks[current].add_import("ecma:array", "isArray");
    let cast_bool = chunks[current].add_import("wasm:js-boolean", "cast");
    let obj_keys = chunks[current].add_import("ecma:object", "keys");
    let obj_get = chunks[current].add_import("ecma:object", "get");
    let str_len = chunks[current].add_import("ecma:string", "length");
    let str_repeat = chunks[current].add_import("ecma:string", "repeat");

    chunks[current].emit_op_u16(Op::LOCAL_SET, items, line);

    // A single object is a stream of one. Every shape below iterates.
    chunks[current].emit_op_u16(Op::LOCAL_GET, items, line);
    chunks[current].emit_call(is_array, 1, line);
    chunks[current].emit_call(cast_bool, 1, line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_if(line);
    // ⛔`push` ANSWERS THE NEW LENGTH, not the array. Storing its result made
    // `items` the number 1, so `length` read 0 and every single-object format
    // rendered empty while the same call on an array worked.
    chunks[current].emit_call(arr_new, 0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, wrapped, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, wrapped, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, items, line);
    chunks[current].emit_call(arr_push, 2, line);
    chunks[current].emit_op(Op::DROP, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, wrapped, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, items, line);
    chunks[current].emit_end(line);

    chunks[current].emit_string_const("", line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);

    // `Format-Table`'s header names the FIRST object's properties, because a
    // table has one shape for the whole stream.
    if mode == FormatMode::Table {
        chunks[current].emit_op_u16(Op::LOCAL_GET, items, line);
        chunks[current].emit_call(arr_len, 1, line);
        chunks[current].emit_i32_const(0, line);
        chunks[current].emit_op(Op::I32_GT_S, line);
        chunks[current].emit_if(line);

        chunks[current].emit_op_u16(Op::LOCAL_GET, items, line);
        chunks[current].emit_i32_const(0, line);
        chunks[current].emit_call(arr_get, 2, line);
        chunks[current].emit_call(obj_keys, 1, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, keys, line);

        for underline in [false, true] {
            chunks[current].emit_i32_const(0, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, key_cursor, line);
            chunks[current].emit_block(line);
            chunks[current].emit_loop_s(line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, key_cursor, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, keys, line);
            chunks[current].emit_call(arr_len, 1, line);
            chunks[current].emit_op(Op::I32_GE_S, line);
            chunks[current].emit_br_if(1, line);

            chunks[current].emit_op_u16(Op::LOCAL_GET, keys, line);
            chunks[current].emit_op_u16(Op::LOCAL_GET, key_cursor, line);
            chunks[current].emit_call(arr_get, 2, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, key, line);

            chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
            if underline {
                // `----` under each heading, as wide as the heading itself.
                chunks[current].emit_string_const("-", line);
                chunks[current].emit_op_u16(Op::LOCAL_GET, key, line);
                chunks[current].emit_call(str_len, 1, line);
                chunks[current].emit_call(str_repeat, 2, line);
            } else {
                chunks[current].emit_op_u16(Op::LOCAL_GET, key, line);
            }
            vybe_compiler::primitives::ops::emit_dyn_add(&mut chunks[current], line);
            chunks[current].emit_string_const("  ", line);
            vybe_compiler::primitives::ops::emit_dyn_add(&mut chunks[current], line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);

            chunks[current].emit_op_u16(Op::LOCAL_GET, key_cursor, line);
            chunks[current].emit_i32_const(1, line);
            chunks[current].emit_op(Op::I32_ADD, line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, key_cursor, line);
            chunks[current].emit_br(0, line);
            chunks[current].emit_end(line);
            chunks[current].emit_end(line);

            chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
            chunks[current].emit_string_const("\n", line);
            vybe_compiler::primitives::ops::emit_dyn_add(&mut chunks[current], line);
            chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);
        }
        chunks[current].emit_end(line);
    }

    chunks[current].emit_i32_const(0, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, cursor, line);

    chunks[current].emit_block(line);
    chunks[current].emit_loop_s(line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, cursor, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, items, line);
    chunks[current].emit_call(arr_len, 1, line);
    chunks[current].emit_op(Op::I32_GE_S, line);
    chunks[current].emit_br_if(1, line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, items, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, cursor, line);
    chunks[current].emit_call(arr_get, 2, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, item, line);

    if mode == FormatMode::Wide {
        // `Format-Wide` renders the VALUE, not its members.
        chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, item, line);
        emit_to_display(chunks, current, line);
        vybe_compiler::primitives::ops::emit_dyn_add(&mut chunks[current], line);
        chunks[current].emit_string_const("\n", line);
        vybe_compiler::primitives::ops::emit_dyn_add(&mut chunks[current], line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);
    } else {
        chunks[current].emit_op_u16(Op::LOCAL_GET, item, line);
        chunks[current].emit_call(obj_keys, 1, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, keys, line);

        chunks[current].emit_i32_const(0, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, key_cursor, line);
        chunks[current].emit_block(line);
        chunks[current].emit_loop_s(line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, key_cursor, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, keys, line);
        chunks[current].emit_call(arr_len, 1, line);
        chunks[current].emit_op(Op::I32_GE_S, line);
        chunks[current].emit_br_if(1, line);

        chunks[current].emit_op_u16(Op::LOCAL_GET, keys, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, key_cursor, line);
        chunks[current].emit_call(arr_get, 2, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, key, line);

        chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
        if mode == FormatMode::List {
            chunks[current].emit_op_u16(Op::LOCAL_GET, key, line);
            vybe_compiler::primitives::ops::emit_dyn_add(&mut chunks[current], line);
            chunks[current].emit_string_const(" : ", line);
            vybe_compiler::primitives::ops::emit_dyn_add(&mut chunks[current], line);
        }
        chunks[current].emit_op_u16(Op::LOCAL_GET, item, line);
        chunks[current].emit_op_u16(Op::LOCAL_GET, key, line);
        chunks[current].emit_call(obj_get, 2, line);
        emit_to_display(chunks, current, line);
        vybe_compiler::primitives::ops::emit_dyn_add(&mut chunks[current], line);
        chunks[current].emit_string_const(if mode == FormatMode::List { "\n" } else { "  " }, line);
        vybe_compiler::primitives::ops::emit_dyn_add(&mut chunks[current], line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);

        chunks[current].emit_op_u16(Op::LOCAL_GET, key_cursor, line);
        chunks[current].emit_i32_const(1, line);
        chunks[current].emit_op(Op::I32_ADD, line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, key_cursor, line);
        chunks[current].emit_br(0, line);
        chunks[current].emit_end(line);
        chunks[current].emit_end(line);

        // A list separates objects with a blank line; a table ends the row.
        chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
        chunks[current].emit_string_const(if mode == FormatMode::List { "\n" } else { "\n" }, line);
        vybe_compiler::primitives::ops::emit_dyn_add(&mut chunks[current], line);
        chunks[current].emit_op_u16(Op::LOCAL_SET, out, line);
    }

    chunks[current].emit_op_u16(Op::LOCAL_GET, cursor, line);
    chunks[current].emit_i32_const(1, line);
    chunks[current].emit_op(Op::I32_ADD, line);
    chunks[current].emit_op_u16(Op::LOCAL_SET, cursor, line);
    chunks[current].emit_br(0, line);
    chunks[current].emit_end(line);
    chunks[current].emit_end(line);

    chunks[current].emit_op_u16(Op::LOCAL_GET, out, line);
}

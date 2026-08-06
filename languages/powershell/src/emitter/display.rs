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

//! `ConvertTo-Json` / `ConvertFrom-Json`.
//!
//! Both are thin routes to `vybe_compiler::primitives::json`, the shared JSON
//! primitive that go and pascal already reach the same way. Nothing here
//! implements JSON — an emitter arm exists only because the primitive is called
//! directly by language adapters rather than published under a `common:json.*`
//! dispatch key.

use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

/// `$obj | ConvertTo-Json`. Uses the property-aware stringify rather than the
/// bare `ecma:json.stringify` so a `[PSCustomObject]` serializes its own
/// properties instead of rendering as `{}`.
///
/// `-Depth` and `-Compress` arrive as named arguments and are dropped: they
/// change how the text is laid out, not what it contains, and the shared
/// serializer takes neither.
pub fn emit_to_json(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    for _ in 1..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    vybe_compiler::primitives::json::emit_stringify_props(chunks, current, line);
}

/// `ConvertFrom-Json`. The shared parse yields `null` for text that is not
/// JSON instead of throwing.
pub fn emit_from_json(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    for _ in 1..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    vybe_compiler::primitives::json::emit_parse_or_null(chunks, current, line);
}

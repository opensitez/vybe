//! `JSON GENERATE` / `JSON PARSE`.
//!
//! Both are thin routes to `vybe_compiler::primitives::json`, the shared JSON
//! primitive that go, pascal and powershell already reach the same way. Nothing
//! here implements JSON — an emitter arm exists only because the primitive is
//! called directly by language adapters rather than published under a
//! `common:json.*` dispatch key.
//!
//! ⛔ NOT `host:ecma:json:*`. COBOL's operand is a RECORD, and a record
//! serialized through the bare ECMA surface renders as `{}` — the shared
//! stringify is the one that walks declared members. `xml.rs` is reached the
//! other way, through the published `common:xml.parse` row in the profile.

use vybe_runtime::Chunk;
use vybe_runtime::opcode::Op;

/// `JSON GENERATE dst FROM src` — the record's members become the object's
/// properties, so this uses the property-aware stringify.
///
/// Trailing operands (`COUNT IN n`, `NAME … IS …`, `SUPPRESS …`) are consumed
/// by the walker, not here; anything still on the stack is dropped so the arity
/// the profile declares always matches what the chunk leaves behind.
pub fn emit_json_generate(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    for _ in 1..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    vybe_compiler::primitives::json::emit_stringify_props(chunks, current, line);
}

/// `JSON PARSE src INTO dst`. The shared parse yields `null` for text that is
/// not JSON rather than throwing, which is what lets `ON EXCEPTION` be a
/// COBOL-level branch instead of a trap.
pub fn emit_json_parse(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    for _ in 1..argc {
        chunks[current].emit_op(Op::DROP, line);
    }
    vybe_compiler::primitives::json::emit_parse_or_null(chunks, current, line);
}

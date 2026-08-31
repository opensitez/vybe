//! `System.Net.Http.HttpMethod` and
//! `System.Net.Http.Headers.MediaTypeHeaderValue` — bytecode-only.
//!
//! Neither type was in the catalog, so `HttpMethod.Get` read `undefined` and
//! `MediaTypeHeaderValue.Parse(...)` trapped with `undefined is not callable`.
//!
//! Both are value-shaped records: a minted object carrying its members as
//! stamped fields (which a member read finds with no type inference at all)
//! plus a `ToString` bound to a precomputed `__display`.
//!
//! Measured against `/usr/local/share/dotnet/dotnet` (SDK 10):
//!
//! ```text
//!   HttpMethod.Get.Method                              -> "GET"   (ToString too)
//!   MediaTypeHeaderValue.Parse("application/json; charset=utf-8")
//!       .MediaType -> "application/json"   .CharSet -> "utf-8"
//!   Parse("text/html").CharSet             -> null, NOT ""
//!   Parse("application/xml; charset=\"utf-16\"").CharSet -> "\"utf-16\""
//!                                             — the QUOTES ARE KEPT
//!   Parse("text/plain;charset=US-ASCII").ToString()
//!       -> "text/plain; charset=US-ASCII"  — a space is inserted
//! ```

use std::sync::Arc;

use vybe_compiler::primitives::class_slots::{self, Dest, ObjSource, ValueSource};
use vybe_compiler::primitives::instructions::core_wasm;
use vybe_runtime::opcode::Op;
use vybe_runtime::{Chunk, Value};

use super::object_fields::field_slot;

const TYPE_KEY: &str = "__type";
/// The string `ToString` answers, computed once where the value is minted so
/// the bound method is a single field read for both types.
const DISPLAY_KEY: &str = "__display";

fn push_const(chunk: &mut Chunk, val: Value, line: u32) {
    match &val {
        Value::String(s) => chunk.emit_string_const(s, line),
        Value::F64(f) => chunk.emit_f64_const(*f, line),
        Value::I32(i) => chunk.emit_i32_const(*i, line),
        _ => panic!("push_const: no WASM-compliant encoding for {:?}", val),
    }
}

fn struct_set_drop(chunk: &mut Chunk, key: &str, line: u32) {
    class_slots::emit_class_set(
        chunk,
        ObjSource::Stack,
        &field_slot(key),
        ValueSource::Stack,
        line,
    );
}

fn call(chunk: &mut Chunk, module: &str, func: &str, argc: u8, line: u32) {
    let idx = chunk.add_import(module, func);
    chunk.emit_call(idx, argc, line);
}

fn get(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, slot, line);
}

fn set(chunk: &mut Chunk, slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_SET, slot, line);
}

/// A `ToString` that answers the stamped [`DISPLAY_KEY`]. Shared by both types
/// and deduplicated by chunk name.
fn push_display_chunk(chunks: &mut Vec<Chunk>, line: u32) -> usize {
    const NAME: &str = "__dotnet_net_value_tostring";
    if let Some(idx) = chunks.iter().position(|chunk| chunk.name == NAME) {
        return idx;
    }
    let mut method = Chunk::new(NAME);
    method.arity = 1;
    method.local_count = 1;
    method.emit_op_u16(Op::LOCAL_GET, 0, line);
    class_slots::emit_class_get(
        &mut method,
        ObjSource::Stack,
        &field_slot(DISPLAY_KEY),
        Dest::Stack,
        line,
    );
    method.emit_op(Op::RETURN, line);
    chunks.push(method);
    chunks.len() - 1
}

fn bind_display(chunks: &mut Vec<Chunk>, current: usize, obj_slot: u16, line: u32) {
    let method_idx = push_display_chunk(chunks, line);
    vybe_compiler::primitives::object::emit_bind_method(
        &mut chunks[current],
        obj_slot,
        &vybe_ast::protocol_slot_key(vybe_ast::ProtocolSlot::ToString),
        method_idx,
        line,
    );
}

// ── HttpMethod ──────────────────────────────────────────────────────────────

/// Mint an `HttpMethod` from the verb already on the stack.
/// Stack: `[verb]` → `[obj]`.
fn emit_http_method_from_stack(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let obj_slot = {
        let chunk = &mut chunks[current];
        let verb_slot = chunk.alloc_scratch(2);
        let obj_slot = verb_slot + 1;
        set(chunk, verb_slot, line);

        class_slots::emit_class_alloc(chunk, line);
        set(chunk, obj_slot, line);
        get(chunk, obj_slot, line);

        core_wasm::dup(chunk, line);
        push_const(chunk, Value::String(Arc::from("HttpMethod")), line);
        struct_set_drop(chunk, TYPE_KEY, line);
        // `Method` is the public member and `__display` backs `ToString`; .NET
        // answers the same verb through both.
        for key in ["Method", "method", DISPLAY_KEY] {
            core_wasm::dup(chunk, line);
            get(chunk, verb_slot, line);
            struct_set_drop(chunk, key, line);
        }
        chunk.emit_op(Op::DROP, line);
        obj_slot
    };
    bind_display(chunks, current, obj_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, obj_slot, line);
}

/// `HttpMethod.Get` and its siblings — static, zero-arity.
pub fn emit_http_method_verb(
    chunks: &mut Vec<Chunk>,
    current: usize,
    verb: &str,
    argc: u8,
    line: u32,
) {
    {
        let chunk = &mut chunks[current];
        for _ in 0..argc {
            chunk.emit_op(Op::DROP, line);
        }
        push_const(chunk, Value::String(Arc::from(verb)), line);
    }
    emit_http_method_from_stack(chunks, current, line);
}

/// `new HttpMethod("REPORT")` — a custom verb.
pub fn emit_http_method_new(chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) {
    {
        let chunk = &mut chunks[current];
        for _ in 1..argc {
            chunk.emit_op(Op::DROP, line);
        }
        if argc == 0 {
            push_const(chunk, Value::String(Arc::from("GET")), line);
        }
    }
    emit_http_method_from_stack(chunks, current, line);
}

// ── MediaTypeHeaderValue ────────────────────────────────────────────────────

/// Push `substring(str_slot, start_slot, end_slot)`.
fn substring_slots(chunk: &mut Chunk, str_slot: u16, start_slot: u16, end_slot: u16, line: u32) {
    get(chunk, str_slot, line);
    get(chunk, start_slot, line);
    get(chunk, end_slot, line);
    call(chunk, "ecma:string", "substring", 3, line);
}

/// `MediaTypeHeaderValue.Parse(text)`. Stack: `[text]` → `[obj]`.
pub fn emit_media_type_header_parse(chunks: &mut Vec<Chunk>, current: usize, line: u32) {
    let obj_slot = {
        let chunk = &mut chunks[current];
        let text_slot = chunk.alloc_scratch(11);
        let lower_slot = text_slot + 1;
        let semi_slot = text_slot + 2;
        let media_slot = text_slot + 3;
        let charset_slot = text_slot + 4;
        let len_slot = text_slot + 5;
        let start_slot = text_slot + 6;
        let end_slot = text_slot + 7;
        let rest_slot = text_slot + 8;
        let zero_slot = text_slot + 9;
        let obj_slot = text_slot + 10;

        set(chunk, text_slot, line);
        push_const(chunk, Value::I32(0), line);
        set(chunk, zero_slot, line);

        get(chunk, text_slot, line);
        call(chunk, "wasm:js-string", "length", 1, line);
        set(chunk, len_slot, line);

        // The parameter search is case-insensitive (`Charset=` is legal), but
        // the VALUE is taken from the original text so its case survives.
        get(chunk, text_slot, line);
        call(chunk, "ecma:string", "toLowerCase", 1, line);
        set(chunk, lower_slot, line);

        get(chunk, text_slot, line);
        push_const(chunk, Value::String(Arc::from(";")), line);
        call(chunk, "ecma:string", "indexOf", 2, line);
        set(chunk, semi_slot, line);

        // media type = everything before the first `;`, trimmed.
        get(chunk, semi_slot, line);
        push_const(chunk, Value::I32(0), line);
        vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
        chunk.emit_if_value(line);
        get(chunk, text_slot, line);
        chunk.emit_else(line);
        substring_slots(chunk, text_slot, zero_slot, semi_slot, line);
        chunk.emit_end(line);
        call(chunk, "ecma:string", "trim", 1, line);
        set(chunk, media_slot, line);

        // charset = the `charset=` parameter's value, up to the next `;`.
        // ⛔ NOT unquoted: .NET answers `"utf-16"` WITH the quotes.
        get(chunk, lower_slot, line);
        push_const(chunk, Value::String(Arc::from("charset=")), line);
        call(chunk, "ecma:string", "indexOf", 2, line);
        set(chunk, start_slot, line);

        get(chunk, start_slot, line);
        push_const(chunk, Value::I32(0), line);
        vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
        chunk.emit_if(line);
        // Absent — .NET answers NULL here, not an empty string.
        chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line);
        set(chunk, charset_slot, line);
        chunk.emit_else(line);
        get(chunk, start_slot, line);
        push_const(chunk, Value::I32(8), line);
        chunk.emit_op(Op::F64_ADD, line);
        set(chunk, start_slot, line);
        substring_slots(chunk, text_slot, start_slot, len_slot, line);
        set(chunk, rest_slot, line);
        get(chunk, rest_slot, line);
        push_const(chunk, Value::String(Arc::from(";")), line);
        call(chunk, "ecma:string", "indexOf", 2, line);
        set(chunk, end_slot, line);
        get(chunk, end_slot, line);
        push_const(chunk, Value::I32(0), line);
        vybe_compiler::primitives::ops::emit_dyn_lt(chunk, line);
        chunk.emit_if_value(line);
        get(chunk, rest_slot, line);
        chunk.emit_else(line);
        substring_slots(chunk, rest_slot, zero_slot, end_slot, line);
        chunk.emit_end(line);
        call(chunk, "ecma:string", "trim", 1, line);
        set(chunk, charset_slot, line);
        chunk.emit_end(line);

        class_slots::emit_class_alloc(chunk, line);
        set(chunk, obj_slot, line);
        get(chunk, obj_slot, line);

        core_wasm::dup(chunk, line);
        push_const(chunk, Value::String(Arc::from("MediaTypeHeaderValue")), line);
        struct_set_drop(chunk, TYPE_KEY, line);
        for key in ["MediaType", "mediatype"] {
            core_wasm::dup(chunk, line);
            get(chunk, media_slot, line);
            struct_set_drop(chunk, key, line);
        }
        for key in ["CharSet", "charset"] {
            core_wasm::dup(chunk, line);
            get(chunk, charset_slot, line);
            struct_set_drop(chunk, key, line);
        }

        // ⛔ `ToString` is NOT `mediaType + "; charset=" + charSet`: .NET keeps
        // EVERY parameter, in its original spelling, and only normalizes the
        // separator to `"; "`. Measured:
        //
        //     "multipart/form-data; boundary=xyz"   -> unchanged (no charset)
        //     "text/csv; charset=utf-8; boundary=q" -> unchanged (both kept)
        //     "text/x; CHARSET=Latin1"              -> case PRESERVED
        //     "text/plain;charset=US-ASCII"         -> a space is INSERTED
        //
        // So the display is the trimmed input with `";"` re-spaced: collapse
        // `"; "` to `";"` first so an already-spaced header does not end up
        // with two. A `;` inside a quoted parameter value would be re-spaced
        // too — .NET does not do that, and no corpus test reaches it.
        core_wasm::dup(chunk, line);
        get(chunk, text_slot, line);
        call(chunk, "ecma:string", "trim", 1, line);
        push_const(chunk, Value::String(Arc::from("; ")), line);
        push_const(chunk, Value::String(Arc::from(";")), line);
        call(chunk, "ecma:string", "replaceAll", 3, line);
        push_const(chunk, Value::String(Arc::from(";")), line);
        push_const(chunk, Value::String(Arc::from("; ")), line);
        call(chunk, "ecma:string", "replaceAll", 3, line);
        struct_set_drop(chunk, DISPLAY_KEY, line);

        chunk.emit_op(Op::DROP, line);
        obj_slot
    };
    bind_display(chunks, current, obj_slot, line);
    chunks[current].emit_op_u16(Op::LOCAL_GET, obj_slot, line);
}

/// `ToString()` for either type — the precomputed [`DISPLAY_KEY`]. The bound
/// protocol slot answers interpolation; this answers the explicit call.
pub fn emit_display_to_string(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    let chunk = &mut chunks[current];
    for _ in 1..argc {
        chunk.emit_op(Op::DROP, line);
    }
    class_slots::emit_class_get(
        chunk,
        ObjSource::Stack,
        &field_slot(DISPLAY_KEY),
        Dest::Stack,
        line,
    );
}

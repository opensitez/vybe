//! Shared .NET JSON adapters.

use vybe_runtime::Chunk;

fn call_import(
    chunks: &mut [Chunk],
    current: usize,
    module: &str,
    name: &str,
    argc: u8,
    line: u32,
) {
    let idx = chunks[current].add_import(module, name);
    let chunk = &mut chunks[current];
    chunk.emit_call(idx, argc, line);
}

pub fn emit_json_serialize(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    call_import(chunks, current, "ecma:json", "stringify", argc, line);
}

pub fn emit_json_deserialize(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    call_import(chunks, current, "ecma:json", "parse", argc, line);
}

// ── System.Text.Json object-shaped values ───────────────────────────────────
//
// `JsonSerializerOptions` and `DefaultJsonTypeInfoResolver` are plain carriers:
// every member the corpus reaches is a property with a default, so they mint an
// object with those defaults stamped and let ordinary member reads and writes
// do the rest. Defaults measured against .NET SDK 10:
//
//   PropertyNameCaseInsensitive false   WriteIndented false
//   AllowTrailingCommas         false   MaxDepth      0
//   Converters                  []      PropertyNamingPolicy / TypeInfoResolver null
//
// ⛔ `MaxDepth` really is `0` and not `64`: .NET reports the UNSET value here
// and substitutes the effective 64 internally, so a "sensible" default would be
// wrong in exactly the place a test looks.

use std::sync::Arc;
use vybe_compiler::primitives::class_slots::{self, ObjSource, ValueSource};
use vybe_compiler::primitives::instructions::core_wasm;
use vybe_runtime::opcode::Op;
use vybe_runtime::Value;

use super::object_fields::field_slot;

/// One stamped field, from a constant.
fn stamp(chunk: &mut Chunk, key: &str, value: Value, line: u32) {
    core_wasm::dup(chunk, line);
    match value {
        Value::String(s) => chunk.emit_string_const(&s, line),
        Value::F64(n) => chunk.emit_i32_const(n as i32, line),
        Value::Bool(b) => core_wasm::bool_const(chunk, line, b),
        _ => chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line),
    }
    class_slots::emit_class_set(
        chunk,
        ObjSource::Stack,
        &field_slot(key),
        ValueSource::Stack,
        line,
    );
}

/// Stamp a fresh empty array under `key`.
fn stamp_empty_array(chunk: &mut Chunk, key: &str, line: u32) {
    core_wasm::dup(chunk, line);
    let idx = chunk.add_import("ecma:array", "new");
    chunk.emit_call(idx, 0, line);
    class_slots::emit_class_set(
        chunk,
        ObjSource::Stack,
        &field_slot(key),
        ValueSource::Stack,
        line,
    );
}

/// Mint an object with `__type` set and the given constant fields, both the
/// declared spelling and its lowercase twin — a member read canonicalises, and
/// which spelling reaches the field depends on whether the receiver's type was
/// inferred at the site.
fn mint(
    chunks: &mut [Chunk],
    current: usize,
    argc: u8,
    type_name: &str,
    bools: &[&str],
    nums: &[(&str, i32)],
    nulls: &[&str],
    arrays: &[&str],
    line: u32,
) {
    let chunk = &mut chunks[current];
    for _ in 0..argc {
        chunk.emit_op(Op::DROP, line);
    }
    let obj = chunk.alloc_scratch(1);
    class_slots::emit_class_alloc(chunk, line);
    chunk.emit_op_u16(Op::LOCAL_SET, obj, line);
    chunk.emit_op_u16(Op::LOCAL_GET, obj, line);

    stamp(chunk, "__type", Value::String(Arc::from(type_name)), line);
    for key in bools {
        for spelling in [key.to_string(), key.to_lowercase()] {
            stamp(chunk, &spelling, Value::Bool(false), line);
        }
    }
    for (key, value) in nums {
        for spelling in [key.to_string(), key.to_lowercase()] {
            stamp(chunk, &spelling, Value::F64(*value as f64), line);
        }
    }
    for key in nulls {
        for spelling in [key.to_string(), key.to_lowercase()] {
            stamp(chunk, &spelling, Value::Null, line);
        }
    }
    for key in arrays {
        for spelling in [key.to_string(), key.to_lowercase()] {
            stamp_empty_array(chunk, &spelling, line);
        }
    }
    chunk.emit_op(Op::DROP, line);
    chunk.emit_op_u16(Op::LOCAL_GET, obj, line);
}

/// `new JsonSerializerOptions()`.
pub fn emit_serializer_options_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    mint(
        chunks,
        current,
        argc,
        "JsonSerializerOptions",
        &[
            "PropertyNameCaseInsensitive",
            "WriteIndented",
            "AllowTrailingCommas",
            "IgnoreReadOnlyProperties",
            "IgnoreReadOnlyFields",
        ],
        &[("MaxDepth", 0)],
        &[
            "PropertyNamingPolicy",
            "TypeInfoResolver",
            "DictionaryKeyPolicy",
            "Encoder",
            "ReferenceHandler",
        ],
        &["Converters"],
        line,
    );
}

/// `new DefaultJsonTypeInfoResolver()`.
pub fn emit_type_info_resolver_new(chunks: &mut [Chunk], current: usize, argc: u8, line: u32) {
    mint(
        chunks,
        current,
        argc,
        "DefaultJsonTypeInfoResolver",
        &[],
        &[],
        &[],
        &["Modifiers"],
        line,
    );
}

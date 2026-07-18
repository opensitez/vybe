//! Shared reflection substrate for language adapters.
//!
//! This module is intentionally not "JavaScript reflection". Vybe's runtime
//! values are ECMA-shaped enough that `ecma:reflect`, `ecma:object`, and
//! `ecma:value` are the portable primitive operations, but each source language
//! still owns its public API:
//!
//! - JavaScript maps these helpers to `typeof`, `instanceof`, `Reflect.*`, and
//!   `Object.*` with ECMA quirks such as prototypes and property descriptors.
//! - PHP maps them to `gettype`, `is_*`, `get_class`, `is_a`, ReflectionClass,
//!   attributes, visibility filters, and dynamic properties.
//! - Go maps them to `reflect.Type` / `reflect.Value`, Kind, Elem, CanSet,
//!   struct tags, and pointer/ref-aware mutation.
//!
//! The shared contract is the hidden metadata/stamp shape plus bytecode recipes
//! for reading and writing live values. Declaration metadata (fields, methods,
//! attributes/tags) is carried by language/compiler metadata; this module only
//! emits the compatible runtime objects that expose it.

use std::sync::Arc;

use vybe_bytecode::opcode::Op;
use vybe_bytecode::{Chunk, Value};

pub const FIELD_TYPE: &str = "__type";
pub const FIELD_TYPES: &str = "__types";
pub const FIELD_TYPE_ID: &str = "__type_id";
pub const FIELD_TYPE_NAME: &str = "__typename";
pub const FIELD_FIELDS: &str = "__fields";
pub const FIELD_FIELDS_PUBLIC: &str = "__fields_public";
pub const FIELD_METHODS: &str = "__methods";
pub const FIELD_METHODS_PUBLIC: &str = "__methods_public";
pub const FIELD_ATTRIBUTES: &str = "__attributes";
pub const FIELD_TAGS: &str = "__tags";
pub const FIELD_KIND: &str = "__kind";
pub const FIELD_VALUE: &str = "__value";
pub const FIELD_REF: &str = "__ref";

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum ReflectKind {
    Undefined,
    Null,
    Bool,
    Number,
    String,
    Symbol,
    Function,
    Object,
    Array,
    Map,
    Set,
    Struct,
    Class,
    Interface,
    Pointer,
    Slice,
    Channel,
}

impl ReflectKind {
    pub fn as_str(self) -> &'static str {
        match self {
            ReflectKind::Undefined => "undefined",
            ReflectKind::Null => "null",
            ReflectKind::Bool => "bool",
            ReflectKind::Number => "number",
            ReflectKind::String => "string",
            ReflectKind::Symbol => "symbol",
            ReflectKind::Function => "function",
            ReflectKind::Object => "object",
            ReflectKind::Array => "array",
            ReflectKind::Map => "map",
            ReflectKind::Set => "set",
            ReflectKind::Struct => "struct",
            ReflectKind::Class => "class",
            ReflectKind::Interface => "interface",
            ReflectKind::Pointer => "ptr",
            ReflectKind::Slice => "slice",
            ReflectKind::Channel => "chan",
        }
    }
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum Visibility {
    Public,
    Protected,
    Private,
}

impl Visibility {
    pub fn as_str(self) -> &'static str {
        match self {
            Visibility::Public => "public",
            Visibility::Protected => "protected",
            Visibility::Private => "private",
        }
    }
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum ObjectKeysMode {
    Own,
    ForIn,
    Values,
    Entries,
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum ObjectOp {
    Object,
    Assign,
    Freeze,
    FromEntries,
    Create,
    Seal,
    IsFrozen,
    IsSealed,
    Is,
    GetPrototypeOf,
    GetOwnPropertyNames,
    GetOwnPropertyDescriptor,
    GetOwnPropertyDescriptors,
    GetOwnPropertySymbols,
    DefineProperty,
    DefineProperties,
    PreventExtensions,
    IsExtensible,
    SetPrototypeOf,
    GroupBy,
    Delete,
    Get,
    Set,
    TrackKey,
    PropertyIsEnumerable,
    HasOwnProperty,
    IsPrototypeOf,
}

impl ObjectOp {
    fn host_name(self) -> &'static str {
        match self {
            ObjectOp::Object => "Object",
            ObjectOp::Assign => "assign",
            ObjectOp::Freeze => "freeze",
            ObjectOp::FromEntries => "fromEntries",
            ObjectOp::Create => "create",
            ObjectOp::Seal => "seal",
            ObjectOp::IsFrozen => "isFrozen",
            ObjectOp::IsSealed => "isSealed",
            ObjectOp::Is => "is",
            ObjectOp::GetPrototypeOf => "getPrototypeOf",
            ObjectOp::GetOwnPropertyNames => "getOwnPropertyNames",
            ObjectOp::GetOwnPropertyDescriptor => "getOwnPropertyDescriptor",
            ObjectOp::GetOwnPropertyDescriptors => "getOwnPropertyDescriptors",
            ObjectOp::GetOwnPropertySymbols => "getOwnPropertySymbols",
            ObjectOp::DefineProperty => "defineProperty",
            ObjectOp::DefineProperties => "defineProperties",
            ObjectOp::PreventExtensions => "preventExtensions",
            ObjectOp::IsExtensible => "isExtensible",
            ObjectOp::SetPrototypeOf => "setPrototypeOf",
            ObjectOp::GroupBy => "groupBy",
            ObjectOp::Delete => "delete",
            ObjectOp::Get => "get",
            ObjectOp::Set => "set",
            ObjectOp::TrackKey => "trackKey",
            ObjectOp::PropertyIsEnumerable => "propertyIsEnumerable",
            ObjectOp::HasOwnProperty => "hasOwnProperty",
            ObjectOp::IsPrototypeOf => "isPrototypeOf",
        }
    }
}

#[derive(Clone, Copy, Debug, Eq, PartialEq)]
pub enum ReflectOp {
    Get,
    Set,
    Apply,
    Construct,
    DeleteProperty,
    DefineProperty,
    GetOwnPropertyDescriptor,
    GetPrototypeOf,
    Has,
    IsExtensible,
    OwnKeys,
    PreventExtensions,
    SetPrototypeOf,
}

impl ReflectOp {
    fn host_name(self) -> &'static str {
        match self {
            ReflectOp::Get => "get",
            ReflectOp::Set => "set",
            ReflectOp::Apply => "apply",
            ReflectOp::Construct => "construct",
            ReflectOp::DeleteProperty => "deleteProperty",
            ReflectOp::DefineProperty => "defineProperty",
            ReflectOp::GetOwnPropertyDescriptor => "getOwnPropertyDescriptor",
            ReflectOp::GetPrototypeOf => "getPrototypeOf",
            ReflectOp::Has => "has",
            ReflectOp::IsExtensible => "isExtensible",
            ReflectOp::OwnKeys => "ownKeys",
            ReflectOp::PreventExtensions => "preventExtensions",
            ReflectOp::SetPrototypeOf => "setPrototypeOf",
        }
    }
}

fn sconst(chunk: &mut Chunk, s: &str) -> u16 {
    chunk.add_constant(Value::String(Arc::from(s)))
}

/// Stack: `[value] -> [ecma_type_string]`.
pub fn emit_typeof(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_import_call(chunks, current, "ecma:value", "typeof", 1, line);
}

/// Single-chunk variant. Stack: `[value] -> [ecma_type_string]`.
pub fn emit_typeof_in_chunk(chunk: &mut Chunk, line: u32) {
    emit_import_call_in_chunk(chunk, "ecma:value", "typeof", 1, line);
}

/// Stack: `[callable] -> [bool]`.
pub fn emit_is_callable(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_import_call(chunks, current, "ecma:reflect", "isCallable", 1, line);
}

/// Single-chunk variant. Stack: `[callable] -> [bool]`.
pub fn emit_is_callable_in_chunk(chunk: &mut Chunk, line: u32) {
    emit_import_call_in_chunk(chunk, "ecma:reflect", "isCallable", 1, line);
}

/// Stack: `[object, key] -> [value]`.
pub fn emit_get_property(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_import_call(chunks, current, "ecma:reflect", "get", 2, line);
}

/// Single-chunk variant. Stack: `[object, key] -> [value]`.
pub fn emit_get_property_in_chunk(chunk: &mut Chunk, line: u32) {
    emit_import_call_in_chunk(chunk, "ecma:reflect", "get", 2, line);
}

/// Stack: `[object, key, value] -> [bool]`.
pub fn emit_set_property(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_import_call(chunks, current, "ecma:reflect", "set", 3, line);
}

/// Single-chunk variant. Stack: `[object, key, value] -> [bool]`.
pub fn emit_set_property_in_chunk(chunk: &mut Chunk, line: u32) {
    emit_import_call_in_chunk(chunk, "ecma:reflect", "set", 3, line);
}

/// Stack: `[object, key] -> [bool]`.
pub fn emit_has_own(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_import_call(chunks, current, "ecma:object", "hasOwn", 2, line);
}

/// Single-chunk variant. Stack: `[object, key] -> [bool]`.
pub fn emit_has_own_in_chunk(chunk: &mut Chunk, line: u32) {
    emit_import_call_in_chunk(chunk, "ecma:object", "hasOwn", 2, line);
}

/// Stack: `[object, key] -> [bool]`.
pub fn emit_has_in(chunks: &mut [Chunk], current: usize, line: u32) {
    emit_import_call(chunks, current, "ecma:object", "hasIn", 2, line);
}

/// Single-chunk variant. Stack: `[object, key] -> [bool]`.
pub fn emit_has_in_in_chunk(chunk: &mut Chunk, line: u32) {
    emit_import_call_in_chunk(chunk, "ecma:object", "hasIn", 2, line);
}

/// Stack: `[object] -> [array]`, where the chosen mode preserves the language
/// adapter's enumeration rules.
pub fn emit_object_view(chunks: &mut [Chunk], current: usize, mode: ObjectKeysMode, line: u32) {
    let name = match mode {
        ObjectKeysMode::Own => "keys",
        ObjectKeysMode::ForIn => "iterForIn",
        ObjectKeysMode::Values => "values",
        ObjectKeysMode::Entries => "entries",
    };
    emit_import_call(chunks, current, "ecma:object", name, 1, line);
}

/// Generic ECMA object operation routed through the shared reflection substrate.
/// Stack contract is the underlying host operation's contract.
pub fn emit_object_op(chunks: &mut [Chunk], current: usize, op: ObjectOp, argc: u8, line: u32) {
    emit_import_call(chunks, current, "ecma:object", op.host_name(), argc, line);
}

/// Generic ECMA Reflect operation routed through the shared reflection substrate.
/// Stack contract is the underlying host operation's contract.
pub fn emit_reflect_op(chunks: &mut [Chunk], current: usize, op: ReflectOp, argc: u8, line: u32) {
    emit_import_call(chunks, current, "ecma:reflect", op.host_name(), argc, line);
}

/// Single-chunk variant. Stack: `[object] -> [array]`.
pub fn emit_object_view_in_chunk(chunk: &mut Chunk, mode: ObjectKeysMode, line: u32) {
    let name = match mode {
        ObjectKeysMode::Own => "keys",
        ObjectKeysMode::ForIn => "iterForIn",
        ObjectKeysMode::Values => "values",
        ObjectKeysMode::Entries => "entries",
    };
    emit_import_call_in_chunk(chunk, "ecma:object", name, 1, line);
}

/// Stack: `[object, class_name] -> [bool]`.
pub fn emit_instanceof(chunks: &mut [Chunk], current: usize, line: u32) {
    crate::classes::emit_instanceof(chunks, current, line);
}

/// Stack: unchanged. Writes `object.__type = type_name`.
pub fn emit_stamp_type(chunk: &mut Chunk, object_slot: u16, type_name: &str, line: u32) {
    emit_set_slot_string_field(chunk, object_slot, FIELD_TYPE, type_name, line);
}

/// Stack: unchanged. Writes `object.__kind = kind`.
pub fn emit_stamp_kind(chunk: &mut Chunk, object_slot: u16, kind: ReflectKind, line: u32) {
    emit_set_slot_string_field(chunk, object_slot, FIELD_KIND, kind.as_str(), line);
}

/// Stack: unchanged. Writes `object[field] = string_value`.
pub fn emit_set_slot_string_field(
    chunk: &mut Chunk,
    object_slot: u16,
    field: &str,
    value: &str,
    line: u32,
) {
    chunk.emit_op_u16(Op::LOCAL_GET, object_slot, line);
    chunk.emit_string_const(value, line);
    let key = sconst(chunk, field);
    chunk.emit_op_u16(Op::STRUCT_SET, key, line);
    chunk.emit_op(Op::DROP, line);
}

/// Stack: unchanged. Writes `object[field] = local_value`.
pub fn emit_set_slot_field_from_local(
    chunk: &mut Chunk,
    object_slot: u16,
    field: &str,
    value_slot: u16,
    line: u32,
) {
    chunk.emit_op_u16(Op::LOCAL_GET, object_slot, line);
    chunk.emit_op_u16(Op::LOCAL_GET, value_slot, line);
    let key = sconst(chunk, field);
    chunk.emit_op_u16(Op::STRUCT_SET, key, line);
    chunk.emit_op(Op::DROP, line);
}

/// Stack: unchanged. Writes `object[field] = ref_func(function_chunk)`.
pub fn emit_bind_method(
    chunk: &mut Chunk,
    object_slot: u16,
    field: &str,
    function_chunk: usize,
    line: u32,
) {
    chunk.emit_op_u16(Op::LOCAL_GET, object_slot, line);
    chunk.emit_op_u16(Op::REF_FUNC, function_chunk as u16, line);
    chunk.emit(0, line);
    let key = sconst(chunk, field);
    chunk.emit_op_u16(Op::STRUCT_SET, key, line);
    chunk.emit_op(Op::DROP, line);
}

/// Create a reflection-shaped object, stamp its type, copy local-backed fields,
/// bind method functions, and leave the object on the stack.
///
/// Stack: unchanged before fields/methods; result `[object]`.
pub fn emit_new_reflection_object(
    chunk: &mut Chunk,
    object_slot: u16,
    type_name: &str,
    fields: &[(&str, u16)],
    methods: &[(&str, usize)],
    line: u32,
) {
    chunk.emit_op_u16(Op::STRUCT_NEW, 0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, object_slot, line);
    emit_stamp_type(chunk, object_slot, type_name, line);
    for (field, value_slot) in fields {
        emit_set_slot_field_from_local(chunk, object_slot, field, *value_slot, line);
    }
    for (method, function_chunk) in methods {
        emit_bind_method(chunk, object_slot, method, *function_chunk, line);
    }
    chunk.emit_op_u16(Op::LOCAL_GET, object_slot, line);
}

/// Create `{ __type, __typename, __kind, __fields, __methods, __attributes }`
/// from local-backed metadata and leave it on the stack. Language adapters may
/// stamp extra fields afterward for public API quirks.
pub fn emit_type_descriptor(
    chunk: &mut Chunk,
    descriptor_slot: u16,
    type_name_slot: u16,
    kind: ReflectKind,
    fields_slot: u16,
    methods_slot: u16,
    attributes_slot: u16,
    line: u32,
) {
    emit_new_reflection_object(
        chunk,
        descriptor_slot,
        "ReflectionType",
        &[
            (FIELD_TYPE_NAME, type_name_slot),
            (FIELD_FIELDS, fields_slot),
            (FIELD_METHODS, methods_slot),
            (FIELD_ATTRIBUTES, attributes_slot),
        ],
        &[],
        line,
    );
    chunk.emit_op(Op::DROP, line);
    emit_stamp_kind(chunk, descriptor_slot, kind, line);
    chunk.emit_op_u16(Op::LOCAL_GET, descriptor_slot, line);
}

/// Create `{ __type, __value, __typename, __kind, __ref }` and leave it on the
/// stack. `ref_slot` should contain null when the value is not settable.
pub fn emit_value_descriptor(
    chunk: &mut Chunk,
    descriptor_slot: u16,
    value_slot: u16,
    type_name_slot: u16,
    kind: ReflectKind,
    ref_slot: u16,
    line: u32,
) {
    emit_new_reflection_object(
        chunk,
        descriptor_slot,
        "ReflectionValue",
        &[
            (FIELD_VALUE, value_slot),
            (FIELD_TYPE_NAME, type_name_slot),
            (FIELD_REF, ref_slot),
        ],
        &[],
        line,
    );
    chunk.emit_op(Op::DROP, line);
    emit_stamp_kind(chunk, descriptor_slot, kind, line);
    chunk.emit_op_u16(Op::LOCAL_GET, descriptor_slot, line);
}

fn emit_import_call(
    chunks: &mut [Chunk],
    current: usize,
    module: &str,
    name: &str,
    argc: u8,
    line: u32,
) {
    let idx = chunks[current].add_import(module, name);
    chunks[current].emit_call(idx, argc, line);
}

fn emit_import_call_in_chunk(chunk: &mut Chunk, module: &str, name: &str, argc: u8, line: u32) {
    let idx = chunk.add_import(module, name);
    chunk.emit_call(idx, argc, line);
}

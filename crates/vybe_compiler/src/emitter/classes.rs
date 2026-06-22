//! Class compilation helpers — shared bytecode patterns for classes across all languages.
//!
//! Every Vybe language (VB, JS, C#, Python) compiles classes to the same constructor
//! pattern. This module provides the common emit sequences so all compilers produce
//! compatible bytecode regardless of source language syntax.
//!
//! Cross-language compatibility:
//! - Python `__str__` and JS `toString()` resolve to the same method
//! - VB `Shared`, JS `static`, C# `static`, Python `@staticmethod` all attach to constructor
//! - All languages use `__get_`/`__set_` prefixed closures for property accessors
//! - `set_type_id` + TypeEntry registration is identical everywhere
//!
//! ## Stack discipline
//!
//! These helpers only cover patterns where the stack state is fully determined —
//! no compiler callbacks needed. Field initialization (where a language-specific
//! expression must be compiled between pushing `this` and calling `struct_set`)
//! is left to each compiler. The `struct_set` opcode expects `[obj, val]` on stack
//! and leaves `[val]` — callers must `drop` if they don't need the result.

use std::sync::Arc;
use vybe_bytecode::chunk::TypeEntry;
use vybe_bytecode::opcode::Op;
use vybe_bytecode::{Chunk, Value};

// ── Object creation ─────────────────────────────────────────────────────

/// Create a new empty object and stamp it with type info.
/// Emits: struct_new 0 → local, __type string stamp, __control_name stamp,
/// set_type_id via __tid_ global.
///
/// Stack: unchanged (object stored in this_slot)
///
/// `__control_name` is set to the lowercased class name. For form classes
/// (a user `Class Form1` in any framework — WinForms, MAUI, etc.) this is
/// the key the GUI host's property registry uses, so `Me.Text = "X"` ends
/// up under `("form1", "text")` and `gui.get_property("form1", "text")`
/// reflects the assignment. For non-form classes the field is dead metadata
/// that nothing reads. Stamping it unconditionally keeps the compiler and
/// the resolver from having to detect "is this class a form?" — the same
/// canonical AST and bytecode shape works for both.
pub fn emit_new_typed_object(chunk: &mut Chunk, this_slot: u16, class_name: &str, line: u32) {
    // Create empty object → store in this_slot
    chunk.emit_op_u16(Op::STRUCT_NEW, 0, line);
    chunk.emit_op_u16(Op::LOCAL_SET, this_slot, line);
    chunk.emit_op(Op::DROP, line);

    // Stamp __type string (untyped fallback for typeof/instanceof)
    // struct_set expects [obj, val] → leaves [val]
    chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
    chunk.emit_string_const(class_name, line);
    let type_key = chunk.add_constant(Value::String(Arc::from("__type")));
    chunk.emit_op_u16(Op::STRUCT_SET, type_key, line);
    chunk.emit_op(Op::DROP, line);

    // Stamp __control_name = lowercased class name (canonical control identity).
    chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
    chunk.emit_string_const(&class_name.to_lowercase(), line);
    let cname_key = chunk.add_constant(Value::String(Arc::from("__control_name")));
    chunk.emit_op_u16(Op::STRUCT_SET, cname_key, line);
    chunk.emit_op(Op::DROP, line);

    // Stamp WASM GC type_id via __tid_ global. The caller has already
    // canonicalised `class_name` per the source language's case-
    // sensitivity, and `register_type` stored the type under that
    // same name — VM `load_type_table` populates `__tid_<canon>`,
    // which we look up verbatim here.
    let tid_name = chunk.add_constant(Value::String(Arc::from(
        format!("__tid_{}", class_name).as_str(),
    )));
    chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
    chunk.emit_op_u16(Op::GLOBAL_GET, tid_name, line);
    { let tid_key = chunk.add_constant(Value::String(Arc::from("__type_id"))); chunk.emit_op_u16(Op::STRUCT_SET, tid_key, line); }
    chunk.emit_op(Op::DROP, line);
}

/// Re-stamp type identity on an EXISTING object — a child constructor
/// receives `this` from the parent ctor call carrying the PARENT's
/// identity, so the child must overwrite `__type` and the WASM GC
/// type_id with its own (otherwise instanceof/REF_TEST and
/// constructorOf resolve to the parent class). Same stamps as
/// `emit_new_typed_object` minus the allocation. `class_name` must be
/// canonicalised like there.
pub fn emit_retype_object(chunk: &mut Chunk, this_slot: u16, class_name: &str, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
    chunk.emit_string_const(class_name, line);
    let type_key = chunk.add_constant(Value::String(Arc::from("__type")));
    chunk.emit_op_u16(Op::STRUCT_SET, type_key, line);
    chunk.emit_op(Op::DROP, line);

    let tid_name = chunk.add_constant(Value::String(Arc::from(
        format!("__tid_{}", class_name).as_str(),
    )));
    chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
    chunk.emit_op_u16(Op::GLOBAL_GET, tid_name, line);
    { let tid_key = chunk.add_constant(Value::String(Arc::from("__type_id"))); chunk.emit_op_u16(Op::STRUCT_SET, tid_key, line); }
    chunk.emit_op(Op::DROP, line);
}

/// Stamp `class_name` into `this.__types` array for cross-language instanceof.
/// If `__types` is null/missing, creates an empty array first, then pushes the name.
/// Called once per class in the inheritance chain (child calls after parent constructor).
///
/// Bytecode stack trace:
/// ```text
/// local_get this          // [this]
/// dup                     // [this, this]
/// struct_get "__types"    // [this, types_or_null]
/// dup                     // [this, types_or_null, types_or_null]
/// ref_is_null             // [this, types_or_null, i32]
/// i32.eqz; br_if 0        // [this, types_or_null]
/// drop                    // [this]
/// array_new 0             // [this, []]
/// skip:                   // [this, array]
/// const "class_name"      // [this, array, "class_name"]
/// array_push              // [this, array_with_name]
/// struct_set "__types"    // [] (stored on this)
/// drop                    // []
/// ```
///
/// Stack: unchanged
pub fn emit_instanceof_chain(
    chunks: &mut [Chunk],
    current: usize,
    this_slot: u16,
    class_name: &str,
    line: u32,
) {
    let types_key = chunks[current].add_constant(Value::String(Arc::from("__types")));

    // Stack: []
    chunks[current].emit_op_u16(Op::LOCAL_GET, this_slot, line); // [this]
    chunks[current].emit_dup(line); // [this, this]
    chunks[current].emit_op_u16(Op::STRUCT_GET, types_key, line); // [this, types_or_null]
    chunks[current].emit_dup(line); // [this, tn, tn]
    chunks[current].emit_op(Op::REF_IS_NULL, line); // [this, tn, bool]
    let init_block = chunks[current].emit_block(line);
    chunks[current].emit_op(Op::I32_EQZ, line);
    chunks[current].emit_br_if(0, line);
    chunks[current].emit_op(Op::DROP, line); // [this] (drop the null)
    crate::emitter::collections::emit_array_new(chunks, current, 0, line); // [this, []]
    chunks[current].emit_end(line);
    chunks[current].patch_block(init_block); // skip lands here; [this, array]

    // Push class_name onto array while preserving array on stack.
    // ecma:array.push is [arr, val] → [new_length], so DUP the array
    // first: [this, array] → [this, array, array] → push → [this, array, len] → drop.
    chunks[current].emit_dup(line); // [this, array, array]
    chunks[current].emit_string_const(class_name, line); // [this, array, array, name]
    crate::emitter::collections::emit_push(chunks, current, line); // [this, array, len]
    chunks[current].emit_op(Op::DROP, line); // [this, array]
    // struct_set: [this, array] → sets this.__types = array, leaves array on stack.
    chunks[current].emit_op_u16(Op::STRUCT_SET, types_key, line); // [array]
    chunks[current].emit_op(Op::DROP, line); // []
}

// ── Method binding ──────────────────────────────────────────────────────

/// Bind an instance method on the object: this.<method_name> = ref_func(chunk_idx).
/// Emits: local_get this → ref_func ci → struct_set key → drop
/// Stack: unchanged
pub fn emit_bind_method(
    chunk: &mut Chunk,
    this_slot: u16,
    method_name: &str,
    method_chunk_idx: usize,
    line: u32,
) {
    chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
    chunk.emit_op_u16(Op::REF_FUNC, method_chunk_idx as u16, line);
    chunk.emit(0, line); // 0 upvalues (upvalue capture is compiler-specific)
    let key = chunk.add_constant(Value::String(Arc::from(method_name)));
    chunk.emit_op_u16(Op::STRUCT_SET, key, line);
    chunk.emit_op(Op::DROP, line);
}

fn emit_stamp_rest_metadata(chunk: &mut Chunk, fixed_count: u8, line: u32) {
    chunk.emit_dup(line);
    chunk.emit_f64_const(fixed_count as f64, line);
    let key = chunk.add_constant(Value::String(Arc::from("__vybe_rest_fixed_arity")));
    chunk.emit_op_u16(Op::STRUCT_SET, key, line);
    chunk.emit_op(Op::DROP, line);
}

pub fn emit_bind_bound_method(
    chunk: &mut Chunk,
    this_slot: u16,
    method_name: &str,
    method_chunk_idx: usize,
    rest_fixed_count: Option<u8>,
    line: u32,
) {
    chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
    chunk.emit_op_u16(Op::REF_FUNC, method_chunk_idx as u16, line);
    chunk.emit(0, line); // 0 upvalues (upvalue capture is compiler-specific)
    chunk.emit_dup(line);
    chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
    let receiver_key = chunk.add_constant(Value::String(Arc::from("__vybe_method_receiver")));
    chunk.emit_op_u16(Op::STRUCT_SET, receiver_key, line);
    chunk.emit_op(Op::DROP, line);
    if let Some(fixed_count) = rest_fixed_count {
        emit_stamp_rest_metadata(chunk, fixed_count, line);
    }
    let key = chunk.add_constant(Value::String(Arc::from(method_name)));
    chunk.emit_op_u16(Op::STRUCT_SET, key, line);
    chunk.emit_op(Op::DROP, line);
}

/// Bind a method AND all its cross-language aliases.
/// This is the primary entry point — ensures a method defined in any language
/// is callable from every other language.
///
/// Example: Python defines `__str__`, this also binds `toString` and `tostring`
/// so JS/VB/C# code can call it transparently.
/// Stack: unchanged
pub fn emit_bind_method_with_aliases(
    chunk: &mut Chunk,
    this_slot: u16,
    method_name: &str,
    method_chunk_idx: usize,
    rest_fixed_count: Option<u8>,
    line: u32,
) {
    // Bind under the original name
    emit_bind_method(chunk, this_slot, method_name, method_chunk_idx, line);
    if let Some(fixed_count) = rest_fixed_count {
        chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
        let key = chunk.add_constant(Value::String(Arc::from(method_name)));
        chunk.emit_op_u16(Op::STRUCT_GET, key, line);
        emit_stamp_rest_metadata(chunk, fixed_count, line);
        chunk.emit_op(Op::DROP, line);
    }
    // Bind under all cross-language aliases
    emit_cross_language_aliases(
        chunk,
        this_slot,
        method_name,
        method_chunk_idx,
        rest_fixed_count,
        line,
    );
}

pub fn emit_bind_bound_method_with_aliases(
    chunk: &mut Chunk,
    this_slot: u16,
    method_name: &str,
    method_chunk_idx: usize,
    rest_fixed_count: Option<u8>,
    line: u32,
) {
    emit_bind_bound_method(
        chunk,
        this_slot,
        method_name,
        method_chunk_idx,
        rest_fixed_count,
        line,
    );
    for &alias in cross_language_aliases(method_name) {
        if alias == method_name {
            continue;
        }
        emit_bind_bound_method(
            chunk,
            this_slot,
            alias,
            method_chunk_idx,
            rest_fixed_count,
            line,
        );
    }
}

/// Return the cross-language alias list for a method name.
/// This is the single source of truth for cross-language method resolution.
/// Returns all equivalent names (including the input name itself).
/// Compilers can filter this list (e.g. skip `__get_`/`__set_` prefixed aliases
/// if the language treats the method as a callable, not a property).
pub fn cross_language_aliases(method_name: &str) -> &'static [&'static str] {
    match method_name {
        // String representation: Python __str__ ↔ JS toString() ↔ VB/C# ToString() ↔ Ruby to_s.
        // Note: __get_tostring removed — ToString is a method, not a property.
        // C#/VB walkers preserve source-case `ToString` as the bound name; include
        // the PascalCase + Ruby spellings so cross-language invocation finds it.
        "__str__" | "tostring" | "toString" | "ToString" | "to_s" | "__toString" => &[
            "__str__",
            "toString",
            "tostring",
            "ToString",
            "to_s",
            "__toString",
        ],

        // Debug representation: Python __repr__
        "__repr__" | "todebugstring" | "toDebugString" => {
            &["__repr__", "toDebugString", "todebugstring"]
        }

        // Length/Count: Python __len__ ↔ JS .length ↔ VB/C# .Count
        "__len__" | "__get_length" | "__get_count" => &["__len__", "__get_length", "__get_count"],

        // Truthiness: Python __bool__ ↔ JS valueOf
        "__bool__" | "valueof" | "valueOf" => &["__bool__", "valueOf", "valueof"],

        // Membership test: Python __contains__ ↔ JS includes() ↔ VB/C# Contains()
        "__contains__" | "contains" | "includes" => &["__contains__", "contains", "includes"],

        // Indexing: Python __getitem__/__setitem__ ↔ Dart operator[]/operator[]=
        "__getitem__" | "operator[]" => &["__getitem__", "operator[]"],
        "__setitem__" | "operator[]=" => &["__setitem__", "operator[]="],

        // Iteration: Python __iter__/__next__ ↔ Dart iterator/moveNext ↔ JS Symbol.iterator
        "__iter__" | "iterator" | "getIterator" => &["__iter__", "iterator", "getIterator"],
        "__next__" | "moveNext" => &["__next__", "moveNext"],

        // Equality: Python __eq__ ↔ Dart operator== ↔ VB/C# Equals()
        "__eq__" | "equals" | "operator==" => &["__eq__", "equals", "operator=="],

        // Hashing: Python __hash__ ↔ VB/C# GetHashCode() ↔ Dart hashCode
        "__hash__" | "gethashcode" | "__get_hashcode" => {
            &["__hash__", "gethashcode", "__get_hashcode"]
        }

        // Comparison: Python __lt__/__gt__/etc ↔ Dart operator</>/ ↔ C# CompareTo
        "__lt__" | "operator<" => &["__lt__", "operator<"],
        "__le__" | "operator<=" => &["__le__", "operator<="],
        "__gt__" | "operator>" => &["__gt__", "operator>"],
        "__ge__" | "operator>=" => &["__ge__", "operator>="],

        // Arithmetic: Python __add__/etc ↔ Dart operator+/etc
        "__add__" | "operator+" => &["__add__", "operator+"],
        "__sub__" | "operator-" => &["__sub__", "operator-"],
        "__mul__" | "operator*" => &["__mul__", "operator*"],
        "__truediv__" | "operator/" => &["__truediv__", "operator/"],
        "__mod__" | "operator%" => &["__mod__", "operator%"],

        // Context manager — Python only, no aliases needed
        "__enter__" | "__exit__" => &[],

        // No aliases for regular method names
        _ => &[],
    }
}

/// Emit cross-language aliases for a method name.
/// Maps between Python dunders, JS camelCase, and VB/C# PascalCase.
///
/// The alias table is the single source of truth for cross-language method resolution.
/// All compilers MUST use this when binding methods so that objects are interoperable.
/// Stack: unchanged
pub fn emit_cross_language_aliases(
    chunk: &mut Chunk,
    this_slot: u16,
    method_name: &str,
    method_chunk_idx: usize,
    rest_fixed_count: Option<u8>,
    line: u32,
) {
    for alias in cross_language_aliases(method_name) {
        if *alias != method_name {
            emit_bind_method(chunk, this_slot, alias, method_chunk_idx, line);
            if let Some(fixed_count) = rest_fixed_count {
                chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
                let key = chunk.add_constant(Value::String(Arc::from(*alias)));
                chunk.emit_op_u16(Op::STRUCT_GET, key, line);
                emit_stamp_rest_metadata(chunk, fixed_count, line);
                chunk.emit_op(Op::DROP, line);
            }
        }
    }
}

// ── Super call (cross-language) ────────────────────────────────────────

/// After calling the parent constructor (result on TOS), store it as `this` and
/// save any parent methods that the child will override.
///
/// The compiler handles the actual call: global_get(parent) → push args → call_ref(argc).
/// This helper stores the result and prepares for child method override.
///
/// Stack before: [parent_return_value]  Stack after: []
pub fn emit_super_call_store_result(
    chunk: &mut Chunk,
    this_slot: u16,
    child_method_names: &[&str],
    line: u32,
) {
    // Store parent-created object as this
    chunk.emit_op_u16(Op::LOCAL_SET, this_slot, line);
    chunk.emit_op(Op::DROP, line);

    // Save parent's methods that child will override (for super.method() calls)
    for method_name in child_method_names {
        emit_save_base_method(chunk, this_slot, method_name, line);
    }
}

// ── Inheritance ─────────────────────────────────────────────────────────

/// Save parent's version of a method as __base_<name> before child override.
/// Used for super()/MyBase/base calls.
/// Emits: local_get this → local_get this → struct_get name → struct_set __base_name → drop
/// Stack: unchanged
pub fn emit_save_base_method(chunk: &mut Chunk, this_slot: u16, method_name: &str, line: u32) {
    let base_name = format!("__base_{}", method_name);
    chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line); // obj for struct_set
    chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line); // obj for struct_get
    let prop_idx = chunk.add_constant(Value::String(Arc::from(method_name)));
    chunk.emit_op_u16(Op::STRUCT_GET, prop_idx, line); // val = this.method (parent version)
    let base_idx = chunk.add_constant(Value::String(Arc::from(base_name.as_str())));
    chunk.emit_op_u16(Op::STRUCT_SET, base_idx, line); // this.__base_method = val
    chunk.emit_op(Op::DROP, line);
}

/// Store parent constructor ref as __super on the instance.
/// Stack: unchanged
pub fn emit_store_super(chunk: &mut Chunk, this_slot: u16, parent_name: &str, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
    let parent_c = chunk.add_constant(Value::String(Arc::from(parent_name)));
    chunk.emit_op_u16(Op::GLOBAL_GET, parent_c, line);
    let super_key = chunk.add_constant(Value::String(Arc::from("__super")));
    chunk.emit_op_u16(Op::STRUCT_SET, super_key, line);
    chunk.emit_op(Op::DROP, line);
}

/// Inherit static methods from parent constructor via Object.assign.
/// Caller must have the constructor on TOS (typically via dup before this call).
/// Stack before: [constructor]  Stack after: [constructor]
pub fn emit_inherit_statics(chunk: &mut Chunk, parent_name: &str, line: u32) {
    chunk.emit_dup(line);
    let parent_c = chunk.add_constant(Value::String(Arc::from(parent_name)));
    chunk.emit_op_u16(Op::GLOBAL_GET, parent_c, line);
    let assign_fn = chunk.add_import("ecma:object", "assign");
    chunk.emit_call(assign_fn, 2, line);
    chunk.emit_op(Op::DROP, line);
}

// ── Static methods ──────────────────────────────────────────────────────

/// Attach a static method to the constructor function object.
/// Same pattern as VB Shared, JS static, C# static, Python @staticmethod.
/// Stack: unchanged (reads constructor from local)
pub fn emit_attach_static_method(
    chunk: &mut Chunk,
    ctor_local: u16,
    method_name: &str,
    method_chunk_idx: usize,
    receiver_slot: Option<u16>,
    rest_fixed_count: Option<u8>,
    line: u32,
) {
    chunk.emit_op_u16(Op::LOCAL_GET, ctor_local, line);
    chunk.emit_op_u16(Op::REF_FUNC, method_chunk_idx as u16, line);
    chunk.emit(0, line);
    if let Some(receiver_slot) = receiver_slot {
        chunk.emit_dup(line);
        chunk.emit_op_u16(Op::LOCAL_GET, receiver_slot, line);
        let receiver_key = chunk.add_constant(Value::String(Arc::from("__vybe_method_receiver")));
        chunk.emit_op_u16(Op::STRUCT_SET, receiver_key, line);
        chunk.emit_op(Op::DROP, line);
    }
    if let Some(fixed_count) = rest_fixed_count {
        emit_stamp_rest_metadata(chunk, fixed_count, line);
    }
    let key = chunk.add_constant(Value::String(Arc::from(method_name)));
    chunk.emit_op_u16(Op::STRUCT_SET, key, line);
    chunk.emit_op(Op::DROP, line);
}

// ── Property accessors ──────────────────────────────────────────────────

/// Bind a property getter as __get_<name> on the instance.
/// The getter_chunk_idx should point to a compiled closure with arity=1 (self/this).
/// Stack: unchanged
pub fn emit_bind_getter(
    chunk: &mut Chunk,
    this_slot: u16,
    prop_name: &str,
    getter_chunk_idx: usize,
    line: u32,
) {
    let get_name = format!("__get_{}", prop_name);
    emit_bind_method(chunk, this_slot, &get_name, getter_chunk_idx, line);
}

/// Bind a property setter as __set_<name> on the instance.
/// The setter_chunk_idx should point to a compiled closure with arity=2 (self/this, value).
/// Stack: unchanged
pub fn emit_bind_setter(
    chunk: &mut Chunk,
    this_slot: u16,
    prop_name: &str,
    setter_chunk_idx: usize,
    line: u32,
) {
    let set_name = format!("__set_{}", prop_name);
    emit_bind_method(chunk, this_slot, &set_name, setter_chunk_idx, line);
}

// ── Constructor return ──────────────────────────────────────────────────

/// Emit return-this at the end of a constructor.
/// Stack: [] → returns this to caller
pub fn emit_constructor_return(chunk: &mut Chunk, this_slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
    chunk.emit_op(Op::RETURN, line);
}

// ── Constructor storage ─────────────────────────────────────────────────

/// Store a constructor function as a local + global variable.
/// Stack: unchanged
pub fn emit_store_constructor(
    chunk: &mut Chunk,
    class_name: &str,
    ctor_chunk_idx: usize,
    local_slot: u16,
    line: u32,
) {
    emit_store_constructor_with_upvalues(
        chunk,
        class_name,
        ctor_chunk_idx,
        local_slot,
        &[],
        false,
        line,
    );
}

/// Store a constructor function with upvalue capture. Used for closure-bound
/// parents (e.g. JS mixin pattern `(Base) => class extends Base`) where the
/// constructor body references variables from an enclosing scope.
///
/// Each upvalue entry is `(is_local, index)` — the same wire format the VM
/// reads after `REF_FUNC`. Pass an empty slice for non-closure constructors.
///
/// `case_sensitive`: when `true` (JS profile), the lowercase alias is NOT
/// emitted — otherwise a `class Range` would overwrite a hoisted
/// `function* range` at runtime, silently draining an empty continuation.
///
/// Stack: unchanged
pub fn emit_store_constructor_with_upvalues(
    chunk: &mut Chunk,
    class_name: &str,
    ctor_chunk_idx: usize,
    local_slot: u16,
    upvalues: &[(bool, u8)],
    case_sensitive: bool,
    line: u32,
) {
    chunk.emit_op_u16(Op::REF_FUNC, ctor_chunk_idx as u16, line);
    chunk.emit(upvalues.len() as u8, line);
    for (is_local, index) in upvalues {
        chunk.emit(if *is_local { 1 } else { 0 }, line);
        chunk.emit(*index, line);
    }
    chunk.emit_op_u16(Op::LOCAL_SET, local_slot, line);
    // Store under original name (case-sensitive lookup)
    let global_name = chunk.add_constant(Value::String(Arc::from(class_name)));
    chunk.emit_op_u16(Op::GLOBAL_SET, global_name, line);
    chunk.emit_op(Op::DROP, line);
    // Also store under lowercase alias for cross-language lookup (VB is case-insensitive).
    // Skip in case-sensitive profiles (JS): a `class Range` must NOT overwrite a hoisted
    // `function* range` — the two names are distinct in a case-sensitive language.
    if !case_sensitive {
        let lower = class_name.to_lowercase();
        if lower != class_name {
            chunk.emit_op_u16(Op::LOCAL_GET, local_slot, line);
            let lower_name = chunk.add_constant(Value::String(Arc::from(lower.as_str())));
            chunk.emit_op_u16(Op::GLOBAL_SET, lower_name, line);
            chunk.emit_op(Op::DROP, line);
        }
    }
}

// ── Field initialization ────────────────────────────────────────────────

/// Set a field on the object to null (pre-declaration / auto-property init).
/// Stack: unchanged
pub fn emit_init_field_null(chunk: &mut Chunk, this_slot: u16, field_name: &str, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
    chunk.emit_op(Op::NULL, line);
    let key = chunk.add_constant(Value::String(Arc::from(field_name)));
    chunk.emit_op_u16(Op::STRUCT_SET, key, line);
    chunk.emit_op(Op::DROP, line);
}

/// Push `this` onto the stack to start a field initialization.
/// Caller compiles the value expression next, then calls `emit_init_field_end`.
/// This wraps the language-specific value-compilation in a compiler_common pattern.
pub fn emit_init_field_start(chunk: &mut Chunk, this_slot: u16, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
}

/// Finish a field initialization started with `emit_init_field_start`.
/// Stack before: [this, value]. Stack after: [].
pub fn emit_init_field_end(chunk: &mut Chunk, field_name: &str, line: u32) {
    let key = chunk.add_constant(Value::String(Arc::from(field_name)));
    chunk.emit_op_u16(Op::STRUCT_SET, key, line);
    chunk.emit_op(Op::DROP, line);
}

/// Get a field value from `this`. Stack before: []. Stack after: [value].
pub fn emit_get_field(chunk: &mut Chunk, this_slot: u16, field_name: &str, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line);
    let key = chunk.add_constant(Value::String(Arc::from(field_name)));
    chunk.emit_op_u16(Op::STRUCT_GET, key, line);
}

/// Set a field value on `this` from a value already on the stack.
/// Stack before: [value]. Stack after: [].
pub fn emit_set_field_from_stack(chunk: &mut Chunk, this_slot: u16, field_name: &str, line: u32) {
    // Need [this, value, value]... actually struct_set expects [obj, val].
    // Caller has [value] — we need to insert this BELOW value on stack.
    // Use a temp local approach: store value, push this, push value, struct_set, drop.
    // Simpler: let the caller use start/end pattern when value isn't pre-computed.
    // For pre-computed value: use a local temp.
    let _ = (chunk, this_slot, field_name, line);
    // This pattern is awkward without a swap opcode.
    // Use emit_init_field_start + emit_init_field_end with value compilation in between.
}

// ── Type registration ───────────────────────────────────────────────────

/// Register a type entry in chunk 0's type table.
pub fn register_type(
    chunks: &mut [Chunk],
    name: &str,
    parent: &str,
    fields: Vec<String>,
    methods: Vec<(String, usize)>,
    is_interface: bool,
    implements: Vec<String>,
    constructor_chunk: Option<usize>,
    field_descriptors: std::collections::HashMap<String, vybe_bytecode::chunk::PropertyDescriptor>,
) {
    // The walker is responsible for case-canonicalising the name per
    // its language's case-sensitivity (`Compiler::canon` lowercases
    // for VB/Pascal/COBOL/PHP, preserves case for JS/TS/Python/C#).
    // No forced lowercasing here — that would silently collide
    // distinct types in case-sensitive languages (`B` and `b`).
    chunks[0].types.push(TypeEntry {
        name: name.to_string(),
        parent: parent.to_string(),
        fields,
        methods,
        is_interface,
        implements,
        constructor_chunk,
        field_descriptors,
    });
}

/// Register an interface/trait/protocol in the type table.
/// Interfaces have no constructor and method entries with chunk_idx=0 (signatures only).
/// This is the same across C# `interface`, VB `Interface`, Dart `abstract class`,
/// Python ABC — different syntax, same TypeEntry shape.
pub fn register_interface(
    chunks: &mut [Chunk],
    name: &str,
    methods: Vec<String>,
    parent_interfaces: Vec<String>,
) {
    // Names arrive pre-canonicalised by the walker — see [`register_type`].
    let method_entries: Vec<(String, usize)> = methods.into_iter().map(|m| (m, 0usize)).collect();
    chunks[0].types.push(TypeEntry {
        name: name.to_string(),
        parent: String::new(),
        fields: Vec::new(),
        methods: method_entries,
        is_interface: true,
        implements: parent_interfaces,
        constructor_chunk: None,
        field_descriptors: std::collections::HashMap::new(),
    });
}

/// Register a class that implements one or more interfaces.
/// This is the standard pattern for C# `: IFoo, IBar`, Dart `implements Foo, Bar`,
/// VB `Implements IFoo`, Python `class Foo(IBar)`.
pub fn register_class_with_interfaces(
    chunks: &mut [Chunk],
    name: &str,
    parent: &str,
    fields: Vec<String>,
    methods: Vec<(String, usize)>,
    implements: Vec<String>,
    constructor_chunk: Option<usize>,
) {
    register_type(
        chunks,
        name,
        parent,
        fields,
        methods,
        false,
        implements,
        constructor_chunk,
        std::collections::HashMap::new(),
    );
}

// ── Super call (cross-language) ────────────────────────────────────────

// ── .NET default constructor: auto-call InitializeComponent ─────────────

/// In .NET, if a class defines `InitializeComponent()` (typical of WinForms
/// designer-generated code) and has no explicit constructor, the default
/// constructor must call `InitializeComponent()` automatically. Both VB and
/// C# follow this convention.
///
/// Emits bytecode equivalent to:
///   Me.InitializeComponent()      ' VB
///   this.InitializeComponent();   // C#
///
/// The `this_slot` is the local variable holding the class instance.
/// Call this AFTER instance methods have been attached to `this` (so that
/// `struct_get "initializecomponent"` finds the method).
pub fn emit_auto_init_component(chunk: &mut Chunk, this_slot: u16, line: u32) {
    emit_auto_init_call(chunk, this_slot, "initializecomponent", line);
}

/// Emit a call to `this.<method_name>()` — generalized auto-init for any
/// method listed in the profile's `auto_init_methods`.  The method name is
/// lowercased for the struct_get lookup (all method keys are stored lowercase).
pub fn emit_auto_init_call(chunk: &mut Chunk, this_slot: u16, method_name: &str, line: u32) {
    chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line); // [this]
    let name_idx = chunk.add_constant(Value::String(Arc::from(method_name.to_lowercase())));
    chunk.emit_op_u16(Op::STRUCT_GET, name_idx, line); // [method_ref]
    chunk.emit_op_u16(Op::LOCAL_GET, this_slot, line); // [method_ref, this]
    chunk.emit_op_u8(Op::CALL_REF, 1, line); // call(1) → [result]
    chunk.emit_op(Op::DROP, line); // []
}

// NOTE: needs_auto_init_component() has moved to type_registry.rs where it
// uses the proper CompileTimeTypes hierarchy instead of string matching.

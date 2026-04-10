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
use vybe_bytecode::{Chunk, Value};
use vybe_bytecode::opcode::Op;
use vybe_bytecode::chunk::TypeEntry;

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
    chunk.emit_op_u16(Op::struct_new, 0, line);
    chunk.emit_op_u16(Op::local_set, this_slot, line);
    chunk.emit_op(Op::drop, line);

    // Stamp __type string (untyped fallback for typeof/instanceof)
    // struct_set expects [obj, val] → leaves [val]
    chunk.emit_op_u16(Op::local_get, this_slot, line);
    let type_str = chunk.add_constant(Value::String(Arc::from(class_name)));
    let type_key = chunk.add_constant(Value::String(Arc::from("__type")));
    chunk.emit_op_u16(Op::r#const, type_str, line);
    chunk.emit_op_u16(Op::struct_set, type_key, line);
    chunk.emit_op(Op::drop, line);

    // Stamp __control_name = lowercased class name (canonical control identity).
    chunk.emit_op_u16(Op::local_get, this_slot, line);
    let cname_str = chunk.add_constant(
        Value::String(Arc::from(class_name.to_lowercase().as_str()))
    );
    let cname_key = chunk.add_constant(Value::String(Arc::from("__control_name")));
    chunk.emit_op_u16(Op::r#const, cname_str, line);
    chunk.emit_op_u16(Op::struct_set, cname_key, line);
    chunk.emit_op(Op::drop, line);

    // Stamp WASM GC type_id via __tid_ global (set at load time by TypeRegistry)
    let tid_name = chunk.add_constant(
        Value::String(Arc::from(format!("__tid_{}", class_name.to_lowercase()).as_str()))
    );
    chunk.emit_op_u16(Op::local_get, this_slot, line);
    chunk.emit_op_u16(Op::global_get, tid_name, line);
    chunk.emit_op(Op::set_type_id, line);
    chunk.emit_op(Op::drop, line);
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
/// ref_is_null             // [this, types_or_null, bool]
/// br_if_false skip        // [this, types_or_null]
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
pub fn emit_instanceof_chain(chunk: &mut Chunk, this_slot: u16, class_name: &str, line: u32) {
    let types_key = chunk.add_constant(Value::String(Arc::from("__types")));

    // [this]
    chunk.emit_op_u16(Op::local_get, this_slot, line);
    // [this, this]
    chunk.emit_op(Op::dup, line);
    // [this, types_or_null]
    chunk.emit_op_u16(Op::struct_get, types_key, line);
    // [this, types_or_null, types_or_null]
    chunk.emit_op(Op::dup, line);
    // [this, types_or_null, bool]
    chunk.emit_op(Op::ref_is_null, line);
    // br_if_false → skip past array creation when __types already exists
    let patch_pos = chunk.code.len();
    chunk.emit_op_u16(Op::br_if_false, 0, line); // placeholder offset
    // [this, null] — null path: drop null and create empty array
    chunk.emit_op(Op::drop, line);
    // [this, []]
    chunk.emit_op_u16(Op::array_new, 0, line);
    // Patch the forward jump to land here
    let offset = (chunk.code.len() as i16) - (patch_pos as i16) - 3;
    chunk.code[patch_pos + 1] = (offset >> 8) as u8;
    chunk.code[patch_pos + 2] = (offset & 0xff) as u8;

    // skip: [this, array]
    let name_const = chunk.add_constant(Value::String(Arc::from(class_name)));
    // [this, array, "class_name"]
    chunk.emit_op_u16(Op::r#const, name_const, line);
    // [this, array_with_name]
    chunk.emit_op(Op::array_push, line);
    // struct_set expects [obj, val] → stores this.__types = array_with_name
    chunk.emit_op_u16(Op::struct_set, types_key, line);
    // drop the result left by struct_set
    chunk.emit_op(Op::drop, line);
}

// ── Method binding ──────────────────────────────────────────────────────

/// Bind an instance method on the object: this.<method_name> = ref_func(chunk_idx).
/// Emits: local_get this → ref_func ci → struct_set key → drop
/// Stack: unchanged
pub fn emit_bind_method(chunk: &mut Chunk, this_slot: u16, method_name: &str, method_chunk_idx: usize, line: u32) {
    chunk.emit_op_u16(Op::local_get, this_slot, line);
    chunk.emit_op_u16(Op::ref_func, method_chunk_idx as u16, line);
    chunk.emit(0, line); // 0 upvalues (upvalue capture is compiler-specific)
    let key = chunk.add_constant(Value::String(Arc::from(method_name)));
    chunk.emit_op_u16(Op::struct_set, key, line);
    chunk.emit_op(Op::drop, line);
}

/// Bind a method AND all its cross-language aliases.
/// This is the primary entry point — ensures a method defined in any language
/// is callable from every other language.
///
/// Example: Python defines `__str__`, this also binds `toString` and `tostring`
/// so JS/VB/C# code can call it transparently.
/// Stack: unchanged
pub fn emit_bind_method_with_aliases(chunk: &mut Chunk, this_slot: u16, method_name: &str, method_chunk_idx: usize, line: u32) {
    // Bind under the original name
    emit_bind_method(chunk, this_slot, method_name, method_chunk_idx, line);
    // Bind under all cross-language aliases
    emit_cross_language_aliases(chunk, this_slot, method_name, method_chunk_idx, line);
}

/// Return the cross-language alias list for a method name.
/// This is the single source of truth for cross-language method resolution.
/// Returns all equivalent names (including the input name itself).
/// Compilers can filter this list (e.g. skip `__get_`/`__set_` prefixed aliases
/// if the language treats the method as a callable, not a property).
pub fn cross_language_aliases(method_name: &str) -> &'static [&'static str] {
    match method_name {
        // String representation: Python __str__ ↔ JS toString() ↔ VB/C# ToString()
        // Note: __get_tostring removed — ToString is a method, not a property.
        "__str__" | "tostring" | "toString" =>
            &["__str__", "toString", "tostring"],

        // Debug representation: Python __repr__
        "__repr__" | "todebugstring" | "toDebugString" =>
            &["__repr__", "toDebugString", "todebugstring"],

        // Length/Count: Python __len__ ↔ JS .length ↔ VB/C# .Count
        "__len__" | "__get_length" | "__get_count" =>
            &["__len__", "__get_length", "__get_count"],

        // Truthiness: Python __bool__ ↔ JS valueOf
        "__bool__" | "valueof" | "valueOf" =>
            &["__bool__", "valueOf", "valueof"],

        // Membership test: Python __contains__ ↔ JS includes() ↔ VB/C# Contains()
        "__contains__" | "contains" | "includes" =>
            &["__contains__", "contains", "includes"],

        // Indexing: Python __getitem__/__setitem__ ↔ Dart operator[]/operator[]=
        "__getitem__" | "operator[]" => &["__getitem__", "operator[]"],
        "__setitem__" | "operator[]=" => &["__setitem__", "operator[]="],

        // Iteration: Python __iter__/__next__ ↔ Dart iterator/moveNext ↔ JS Symbol.iterator
        "__iter__" | "iterator" | "getIterator" =>
            &["__iter__", "iterator", "getIterator"],
        "__next__" | "moveNext" =>
            &["__next__", "moveNext"],

        // Equality: Python __eq__ ↔ Dart operator== ↔ VB/C# Equals()
        "__eq__" | "equals" | "operator==" =>
            &["__eq__", "equals", "operator=="],

        // Hashing: Python __hash__ ↔ VB/C# GetHashCode() ↔ Dart hashCode
        "__hash__" | "gethashcode" | "__get_hashcode" =>
            &["__hash__", "gethashcode", "__get_hashcode"],

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
pub fn emit_cross_language_aliases(chunk: &mut Chunk, this_slot: u16, method_name: &str, method_chunk_idx: usize, line: u32) {
    for alias in cross_language_aliases(method_name) {
        if *alias != method_name {
            emit_bind_method(chunk, this_slot, alias, method_chunk_idx, line);
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
    chunk.emit_op_u16(Op::local_set, this_slot, line);
    chunk.emit_op(Op::drop, line);

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
    chunk.emit_op_u16(Op::local_get, this_slot, line);  // obj for struct_set
    chunk.emit_op_u16(Op::local_get, this_slot, line);  // obj for struct_get
    let prop_idx = chunk.add_constant(Value::String(Arc::from(method_name)));
    chunk.emit_op_u16(Op::struct_get, prop_idx, line);   // val = this.method (parent version)
    let base_idx = chunk.add_constant(Value::String(Arc::from(base_name.as_str())));
    chunk.emit_op_u16(Op::struct_set, base_idx, line);   // this.__base_method = val
    chunk.emit_op(Op::drop, line);
}

/// Store parent constructor ref as __super on the instance.
/// Stack: unchanged
pub fn emit_store_super(chunk: &mut Chunk, this_slot: u16, parent_name: &str, line: u32) {
    chunk.emit_op_u16(Op::local_get, this_slot, line);
    let parent_c = chunk.add_constant(Value::String(Arc::from(parent_name)));
    chunk.emit_op_u16(Op::global_get, parent_c, line);
    let super_key = chunk.add_constant(Value::String(Arc::from("__super")));
    chunk.emit_op_u16(Op::struct_set, super_key, line);
    chunk.emit_op(Op::drop, line);
}

/// Inherit static methods from parent constructor via Object.assign.
/// Caller must have the constructor on TOS (typically via dup before this call).
/// Stack before: [constructor]  Stack after: [constructor]
pub fn emit_inherit_statics(chunk: &mut Chunk, parent_name: &str, line: u32) {
    chunk.emit_op(Op::dup, line);
    let parent_c = chunk.add_constant(Value::String(Arc::from(parent_name)));
    chunk.emit_op_u16(Op::global_get, parent_c, line);
    let assign_fn = chunk.add_import("vybe:object", "assign");
    chunk.emit_op_u16(Op::call_import, assign_fn, line);
    chunk.emit(2, line);
    chunk.emit_op(Op::drop, line);
}

// ── Static methods ──────────────────────────────────────────────────────

/// Attach a static method to the constructor function object.
/// Same pattern as VB Shared, JS static, C# static, Python @staticmethod.
/// Stack: unchanged (reads constructor from local)
pub fn emit_attach_static_method(chunk: &mut Chunk, ctor_local: u16, method_name: &str, method_chunk_idx: usize, line: u32) {
    chunk.emit_op_u16(Op::local_get, ctor_local, line);
    chunk.emit_op_u16(Op::ref_func, method_chunk_idx as u16, line);
    chunk.emit(0, line);
    let key = chunk.add_constant(Value::String(Arc::from(method_name)));
    chunk.emit_op_u16(Op::struct_set, key, line);
    chunk.emit_op(Op::drop, line);
}

// ── Property accessors ──────────────────────────────────────────────────

/// Bind a property getter as __get_<name> on the instance.
/// The getter_chunk_idx should point to a compiled closure with arity=1 (self/this).
/// Stack: unchanged
pub fn emit_bind_getter(chunk: &mut Chunk, this_slot: u16, prop_name: &str, getter_chunk_idx: usize, line: u32) {
    let get_name = format!("__get_{}", prop_name);
    emit_bind_method(chunk, this_slot, &get_name, getter_chunk_idx, line);
}

/// Bind a property setter as __set_<name> on the instance.
/// The setter_chunk_idx should point to a compiled closure with arity=2 (self/this, value).
/// Stack: unchanged
pub fn emit_bind_setter(chunk: &mut Chunk, this_slot: u16, prop_name: &str, setter_chunk_idx: usize, line: u32) {
    let set_name = format!("__set_{}", prop_name);
    emit_bind_method(chunk, this_slot, &set_name, setter_chunk_idx, line);
}

// ── Constructor return ──────────────────────────────────────────────────

/// Emit return-this at the end of a constructor.
/// Stack: [] → returns this to caller
pub fn emit_constructor_return(chunk: &mut Chunk, this_slot: u16, line: u32) {
    chunk.emit_op_u16(Op::local_get, this_slot, line);
    chunk.emit_op(Op::r#return, line);
}

// ── Constructor storage ─────────────────────────────────────────────────

/// Store a constructor function as a local + global variable.
/// Stack: unchanged
pub fn emit_store_constructor(chunk: &mut Chunk, class_name: &str, ctor_chunk_idx: usize, local_slot: u16, line: u32) {
    chunk.emit_op_u16(Op::ref_func, ctor_chunk_idx as u16, line);
    chunk.emit(0, line);
    chunk.emit_op_u16(Op::local_set, local_slot, line);
    // Store under original name (case-sensitive lookup)
    let global_name = chunk.add_constant(Value::String(Arc::from(class_name)));
    chunk.emit_op_u16(Op::global_set, global_name, line);
    chunk.emit_op(Op::drop, line);
    // Also store under lowercase alias for cross-language lookup (VB is case-insensitive)
    let lower = class_name.to_lowercase();
    if lower != class_name {
        chunk.emit_op_u16(Op::local_get, local_slot, line);
        let lower_name = chunk.add_constant(Value::String(Arc::from(lower.as_str())));
        chunk.emit_op_u16(Op::global_set, lower_name, line);
        chunk.emit_op(Op::drop, line);
    }
}

// ── Field initialization ────────────────────────────────────────────────

/// Set a field on the object to null (pre-declaration / auto-property init).
/// Stack: unchanged
pub fn emit_init_field_null(chunk: &mut Chunk, this_slot: u16, field_name: &str, line: u32) {
    chunk.emit_op_u16(Op::local_get, this_slot, line);
    chunk.emit_op(Op::null, line);
    let key = chunk.add_constant(Value::String(Arc::from(field_name)));
    chunk.emit_op_u16(Op::struct_set, key, line);
    chunk.emit_op(Op::drop, line);
}

/// Push `this` onto the stack to start a field initialization.
/// Caller compiles the value expression next, then calls `emit_init_field_end`.
/// This wraps the language-specific value-compilation in a compiler_common pattern.
pub fn emit_init_field_start(chunk: &mut Chunk, this_slot: u16, line: u32) {
    chunk.emit_op_u16(Op::local_get, this_slot, line);
}

/// Finish a field initialization started with `emit_init_field_start`.
/// Stack before: [this, value]. Stack after: [].
pub fn emit_init_field_end(chunk: &mut Chunk, field_name: &str, line: u32) {
    let key = chunk.add_constant(Value::String(Arc::from(field_name)));
    chunk.emit_op_u16(Op::struct_set, key, line);
    chunk.emit_op(Op::drop, line);
}

/// Get a field value from `this`. Stack before: []. Stack after: [value].
pub fn emit_get_field(chunk: &mut Chunk, this_slot: u16, field_name: &str, line: u32) {
    chunk.emit_op_u16(Op::local_get, this_slot, line);
    let key = chunk.add_constant(Value::String(Arc::from(field_name)));
    chunk.emit_op_u16(Op::struct_get, key, line);
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
) {
    chunks[0].types.push(TypeEntry {
        name: name.to_lowercase(),
        parent: parent.to_string(),
        fields,
        methods,
        is_interface,
        implements,
        constructor_chunk,
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
    let method_entries: Vec<(String, usize)> = methods.iter()
        .map(|m| (m.to_lowercase(), 0usize))
        .collect();
    chunks[0].types.push(TypeEntry {
        name: name.to_lowercase(),
        parent: String::new(),
        fields: Vec::new(),
        methods: method_entries,
        is_interface: true,
        implements: parent_interfaces.iter().map(|s| s.to_lowercase()).collect(),
        constructor_chunk: None,
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
    register_type(chunks, name, parent, fields, methods, false, implements, constructor_chunk);
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
    // Me.InitializeComponent() → struct_get + call with this as arg
    chunk.emit_op_u16(Op::local_get, this_slot, line);   // [this]
    let name_idx = chunk.add_constant(Value::String(Arc::from("initializecomponent")));
    chunk.emit_op_u16(Op::struct_get, name_idx, line);    // [method_ref]
    chunk.emit_op_u16(Op::local_get, this_slot, line);    // [method_ref, this]
    chunk.emit_op_u8(Op::call, 1, line);                  // call(1) → [result]
    chunk.emit_op(Op::drop, line);                        // []
}

// NOTE: needs_auto_init_component() has moved to type_registry.rs where it
// uses the proper CompileTimeTypes hierarchy instead of string matching.

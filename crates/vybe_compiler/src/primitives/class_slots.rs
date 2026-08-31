//! The class-model owner API.
//!
//! Every name-keyed access to an object's storage goes through here. Callers
//! name a `ClassSlot`; they never spell a storage key and never emit a struct
//! op themselves. That is what lets M6 change `STRUCT_GET 0 <str_const>` into
//! `STRUCT_GET <typeidx> <field_index>` by editing this file alone.
//!
//! The parameter set is not a design preference — each one is a call shape that
//! exists in the tree today and currently hand-rolls its own adapter:
//!
//! - `ObjSource`   — the object is on the STACK in ~2/3 of wrappers, but dart,
//!                   jvm and lua are majority local-slot. One-way costs the
//!                   other group a pointless round trip.
//! - `ValueSource` — the object-source x value-source matrix is fully
//!                   populated; `adodb_adapter.rs` hand-writes four of its six
//!                   cells as four separately named helpers.
//! - `Dest`        — 17 wrappers already carry an output-destination parameter
//!                   and hand-write the trailing `LOCAL_SET` without one.
//!
//! There is no keep-or-drop flag: `STRUCT_SET` consumes the object, so "keep"
//! on the stack costs a `dup` and "drop" from a local is meaningless. The owner
//! derives it from `ObjSource`.

use crate::primitives::instructions::host;
use std::sync::Arc;
use vybe_ast::{ExprKind, Expression, ProtocolSlot};
use vybe_runtime::chunk::Chunk;
use vybe_runtime::opcode::Op;
use vybe_runtime::value::Value;

/// Where the object comes from.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum ObjSource {
    /// Already on the stack, on top.
    Stack,
    /// In a local slot; the owner emits the `LOCAL_GET`.
    Local(u16),
}

/// Where a value comes from.
#[derive(Debug, Clone, PartialEq)]
pub enum ValueSource {
    /// Already on the stack, above the object.
    Stack,
    /// In a local slot.
    Local(u16),
    /// A compile-time constant. Folds the `set_field_const`,
    /// `emit_set_field_const_i32`, `set_field_string`, `set_field_bool`,
    /// `set_field_f64` and `set_field_i32` families into one call.
    ///
    /// Constants are typed at the wasm level — there is no generic
    /// push-a-constant opcode — so the owner dispatches to the right
    /// `emit_*_const`. That dispatch existing in ONE place is the point.
    ConstStr(String),
    ConstI32(i32),
    ConstI64(i64),
    ConstF64(f64),
    ConstBool(bool),
    /// A null reference. Distinct from `ConstStr` etc. because `Value::Null`
    /// has no literal encoding — it is `ref.null extern`.
    Null,

    /// A closure over a compiled function — `REF_FUNC idx` plus an upvalue
    /// count. Method binding is what puts methods on a prototype, so this is
    /// the class model rather than a corner of it, and `Value` has no function
    /// variant to carry it as a constant. 12 sites (crates 2 · languages 10),
    /// including all six of php's `finish_*_instance` constructors.
    FuncRef { idx: u16, upvalues: u8 },
}

/// Where a read result goes.
#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub enum Dest {
    /// Left on the stack.
    Stack,
    /// Stored into a local slot; the owner emits the `LOCAL_SET`.
    Local(u16),
}

/// A method identity. NOT a `ClassSlot`: today a method is a funcref in a
/// name-keyed field, but at M7 dispatch moves to `call_with_tag` on the
/// descriptor, where a method is a TAG. Name+arity is already the call-tag key
/// (`dynamic_invoke::invoke_tag_name`), so this is the shape the tag proposal
/// wants and modelling it as a slot would have to be unpicked.
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct MethodRef {
    pub name: String,
    pub argc: u8,
}

impl MethodRef {
    pub fn new(name: impl Into<String>, argc: u8) -> Self {
        Self { name: name.into(), argc }
    }
}

/// WHICH slot on the object.
#[derive(Debug, Clone, PartialEq)]
pub enum ClassSlot {
    /// A declared instance field.
    ///
    /// `class` is REQUIRED: the private-mangled answer depends on the DECLARING
    /// class, and reading the compiler's `current_class` instead is wrong for
    /// any receiver that is not `self`.
    InstanceField { class: Option<String>, field: String },

    /// A static field. A SEPARATE index space from instance fields — they
    /// cannot fold together without colliding once both become field indices.
    StaticField { class: Option<String>, field: String },

    /// A private field, whose storage name is mangled against its declaring
    /// class.
    PrivateField { class: String, field: String },

    /// An accessor pair. Built with `format!("__get_{key}")` at
    /// `emit_helpers.rs:335-336`, `object.rs:399/412` and `dispatch.rs:4180`,
    /// so these spellings are invisible to any literal-based search.
    Getter(String),
    Setter(String),

    /// `obj.prototype` and `obj.__proto__`. Without these the prototype hop has
    /// nothing to convert to and the string-keyed emissions stay forever.
    Prototype,
    ProtoLink,

    /// Engine-internal state: generator/iterator/stream/weakref slots, adapter
    /// buffers. The DOMINANT variant, not the exceptional one.
    ///
    /// Scoped by REPRESENTATION, never by adapter. `__ms_pos`/`__ms_buf`/
    /// `__ms_len` are shared by binary_io, memory_stream and stream_io;
    /// per-adapter scoping would give them three identities and a missing key
    /// reads as `undefined`, so it would fail silently.
    Internal(String),

    /// A slot whose spelling is shared across trees — `__value`, `__keys`,
    /// `__types`. Globally scoped for the same reason as `Internal`, one level
    /// wider: dotnet writes what libc reads.
    Repr(String),

    /// `__type`. Becomes the RTT at M6.
    TypeIdentity,

    /// A bound protocol slot. The abstraction this plan is named after —
    /// languages bind, they don't name.
    Slot(ProtocolSlot),

    /// The key is NOT known at compile time.
    ///
    /// The explicit escape hatch, for the sites where the key's KIND is only
    /// decidable at runtime — `emit_lua_table_write` type-tests the key to pick
    /// between array-index, `ecma:map.set` and `ecma:object.delete`; a dart Map
    /// with a computed key is the same shape. Without this variant those sites
    /// would be silently excluded from the conversion denominator, and an
    /// exception you can grep is safer than one you cannot.
    Dynamic(ValueSource),
}

impl ClassSlot {
    /// An instance field with no declaring class known — the plain canonical
    /// name. Prefer `instance_of` wherever the declaring class IS known.
    pub fn instance(field: impl Into<String>) -> Self {
        ClassSlot::InstanceField { class: None, field: field.into() }
    }

    pub fn instance_of(class: impl Into<String>, field: impl Into<String>) -> Self {
        ClassSlot::InstanceField { class: Some(class.into()), field: field.into() }
    }

    pub fn internal(key: impl Into<String>) -> Self {
        ClassSlot::Internal(key.into())
    }

    pub fn repr(key: impl Into<String>) -> Self {
        ClassSlot::Repr(key.into())
    }
}

/// Resolves a declared member name to its storage name.
///
/// The compiler implements this with the private-mangling rules; emitters that
/// have no `Compiler` in hand use [`PlainNames`]. Keeping it a trait is what
/// lets adapter code share the owner without depending on the compiler.
pub trait SlotNames {
    fn storage_name(&self, class: Option<&str>, field: &str) -> String;
}

/// Identity resolution, for internal and adapter-private state that is never
/// mangled and never case-folded.
pub struct PlainNames;

impl SlotNames for PlainNames {
    fn storage_name(&self, _class: Option<&str>, field: &str) -> String {
        field.to_string()
    }
}

impl SlotNames for crate::Compiler {
    /// The compiler's resolution: private members mangle against their
    /// DECLARING class, everything else takes the canonical spelling.
    fn storage_name(&self, class: Option<&str>, field: &str) -> String {
        match class {
            Some(class) => self.js_member_storage_name_for_class(class, field),
            None => self.js_member_storage_name(field),
        }
    }
}

/// A slot that has already become a storage key.
///
/// Resolution needs `&Compiler` (private mangling) while emission needs
/// `&mut Chunk`, and both live on the same struct — so the key is resolved
/// FIRST and carried here. That is a borrow-checker fact, but it also names the
/// M6 seam precisely: this type becomes `Index(u16)` and nothing else moves.
#[derive(Debug, Clone)]
pub enum ResolvedSlot {
    /// A compile-time key. Becomes a field index at M6.
    Key(String),
    /// A compile-time key ALREADY INTERNED into a chunk's constant pool.
    ///
    /// ⚠ `Chunk::add_constant` does NOT de-duplicate — it pushes
    /// unconditionally. So a key hoisted above a loop and emitted N times must
    /// intern ONCE, or the conversion silently grows the constant pool by N-1
    /// entries per site. That is the shape `let done_key = …` at the top of
    /// `generators.rs` has, and there are ~150 like it across the tree.
    ///
    /// Resolve with [`resolve_interned`] wherever the original code hoisted its
    /// key, and the emitted pool is identical rather than merely equivalent.
    Interned(u16),
    /// ▶▶ **SEAM 3 — an INDEXED field on a registered type.**
    ///
    /// `struct.get/set <typeidx> <fieldidx>` against the instance's real field
    /// storage, which is what WASM GC means by a struct access. This is the
    /// form the whole plan exists to reach; the variants above are the
    /// string-keyed fallback for objects that genuinely have no declared shape.
    ///
    /// ⚠ Only ever produced when the field is found **BY NAME** in the type's
    /// declared field list. A positional guess is silent corruption: python's
    /// `@dataclass` emits `0:__class__ 1:x 2:y 3:label 4:__dataclass_fields__`,
    /// so `x` is at index **1**, and indexing by declaration position would
    /// write it into `__class__`.
    Indexed { typeidx: u16, field: u16 },
    /// A runtime key, already described by its source.
    Dynamic(ValueSource),
}

/// Resolve a slot and intern its key into `chunk` immediately.
///
/// Use this — not [`resolve`] — wherever the original code hoisted the key
/// above its emit sites. Interning once preserves both the constant-pool
/// CONTENTS and its ORDER; resolving lazily at each emit would append a
/// duplicate entry per use, because `add_constant` does not de-duplicate.
pub fn resolve_interned(chunk: &mut Chunk, slot: &ClassSlot, names: &dyn SlotNames) -> ResolvedSlot {
    match storage_key(slot, names) {
        Some(key) => {
            let idx = chunk.add_constant(Value::String(Arc::from(key.as_str())));
            ResolvedSlot::Interned(idx)
        }
        None => {
            let ClassSlot::Dynamic(src) = slot else { unreachable!() };
            ResolvedSlot::Dynamic(src.clone())
        }
    }
}

/// Resolve a slot to its storage key.
pub fn resolve(slot: &ClassSlot, names: &dyn SlotNames) -> ResolvedSlot {
    match storage_key(slot, names) {
        Some(key) => ResolvedSlot::Key(key),
        None => {
            let ClassSlot::Dynamic(src) = slot else { unreachable!() };
            ResolvedSlot::Dynamic(src.clone())
        }
    }
}

/// The storage name of a getter accessor. **The one definition.**
///
/// Nine walker sites across js/kotlin/php built this with `format!("__get_{}")`
/// by hand, plus five more in vb — a second spelling of a key `ClassSlot::Getter`
/// already owns. Today they match and nothing enforces it, which is exactly how
/// `protocol_slot_key` came to have two spellings and cost 24 VB tests.
pub fn getter_name(key: impl AsRef<str>) -> String {
    format!("__get_{}", key.as_ref())
}

/// The storage name of a setter accessor. See [`getter_name`].
pub fn setter_name(key: impl AsRef<str>) -> String {
    format!("__set_{}", key.as_ref())
}

/// The storage key a slot resolves to TODAY.
///
/// Returns `None` for [`ClassSlot::Dynamic`], whose key is a runtime value and
/// is pushed by the emit functions instead.
///
/// This is the single place a slot becomes a name. At M6 it becomes the single
/// place a slot becomes a field INDEX, and no caller changes.
pub fn storage_key(slot: &ClassSlot, names: &dyn SlotNames) -> Option<String> {
    Some(match slot {
        ClassSlot::InstanceField { class, field }
        | ClassSlot::StaticField { class, field } => {
            names.storage_name(class.as_deref(), field)
        }
        ClassSlot::PrivateField { class, field } => {
            names.storage_name(Some(class), field)
        }
        ClassSlot::Getter(key) => getter_name(key),
        ClassSlot::Setter(key) => setter_name(key),
        ClassSlot::Prototype => "prototype".to_string(),
        ClassSlot::ProtoLink => "__proto__".to_string(),
        ClassSlot::Internal(key) | ClassSlot::Repr(key) => key.clone(),
        ClassSlot::TypeIdentity => "__type".to_string(),
        ClassSlot::Slot(slot) => protocol_slot_key(*slot),
        ClassSlot::Dynamic(_) => return None,
    })
}

/// The storage name a bound protocol slot occupies on an object.
///
/// Languages bind a slot; they never name it. This is the one place the
/// binding becomes a spelling.
fn protocol_slot_key(slot: ProtocolSlot) -> String {
    // ⛔ DELEGATES. This used to be `format!("__slot_{slot:?}").to_lowercase()`
    // — a SECOND spelling for a key that already had one, so
    // `ClassSlot::Slot(ToString)` resolved to `__slot_tostring` while every
    // binder and every dispatch site keys `__vybe_slot_<id>`. The read found
    // nothing and answered `undefined`, which is a silent wrong answer rather
    // than a failure: it cost 24 VB tests, all of them `ToString` overrides and
    // TimeSpan operators.
    //
    // `ProtocolSlot::slot_id` is the stable numeric identity dispatch keys on,
    // so the AST's function is the authority and this is not free to have an
    // opinion.
    vybe_ast::protocol_slot_key(slot)
}

// ── The verbs ───────────────────────────────────────────────────────────

/// Push the object described by `obj`, leaving it on top of the stack.
fn push_obj(chunk: &mut Chunk, obj: ObjSource, line: u32) {
    match obj {
        ObjSource::Stack => {}
        ObjSource::Local(slot) => chunk.emit_op_u16(Op::LOCAL_GET, slot, line),
    }
}

/// Push the value described by `val`.
fn push_value(chunk: &mut Chunk, val: &ValueSource, line: u32) {
    match val {
        ValueSource::Stack => {}
        ValueSource::Local(slot) => chunk.emit_op_u16(Op::LOCAL_GET, *slot, line),
        ValueSource::ConstStr(v) => chunk.emit_string_const(v, line),
        ValueSource::ConstI32(v) => chunk.emit_i32_const(*v, line),
        ValueSource::ConstI64(v) => chunk.emit_i64_const(*v, line),
        ValueSource::ConstF64(v) => chunk.emit_f64_const(*v, line),
        ValueSource::ConstBool(v) => chunk.emit_bool_const(*v, line),
        ValueSource::Null => {
            chunk.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, line)
        }
        ValueSource::FuncRef { idx, upvalues } => {
            chunk.emit_op_u16(Op::REF_FUNC, *idx, line);
            chunk.emit(*upvalues, line);
        }
    }
}

/// Push the key for `slot` — a constant for a static slot, the runtime value
/// for a dynamic one.

/// Declare a name-keyed slot's key as a `wasm:string-constants` global import.
///
/// ⛔ THE WASM WRITER NEEDS THE KEY AS A VALUE, NOT A CONSTANT INDEX. A
/// name-keyed access emits `STRUCT_GET/SET` with typeidx 0 — "no static type",
/// a property by NAME — which the writer lowers to `ecma:object.get`/`set`,
/// because core wasm has no by-name property access. That host call needs the
/// name ON THE STACK, and the only way a module names a string is the
/// js-string-builtins constant import. A key used solely as a struct immediate
/// never got one, so the lowering had nothing to push and fell back to a typed
/// `struct.get 0 0` — what V8 rejects as `invalid field index: 0`.
///
/// ⛔ AND IT HAS TO BE DECLARED HERE, NOT IN THE WRITER. Global IMPORTS precede
/// defined globals, and the writer uses a `GLOBAL_GET`/`GLOBAL_SET` immediate
/// directly as the wasm global index. Adding the names writer-side shifts every
/// defined global out from under immediates the compiler already numbered —
/// measured: V8 answered `immutable global #12 cannot be assigned`. Declared
/// here, the name is in `global_imports` before `normalize_global_table` runs,
/// so both numberings come from one list.
fn declare_key_string(chunk: &mut Chunk, name: &str) {
    chunk.add_global_import(vybe_runtime::chunk::STRING_CONSTANTS_MODULE, name);
}

fn push_key(chunk: &mut Chunk, slot: &ResolvedSlot, line: u32) -> Option<u16> {
    match slot {
        ResolvedSlot::Key(name) => {
            declare_key_string(chunk, name);
            Some(chunk.add_constant(Value::String(Arc::from(name.as_str()))))
        }
        ResolvedSlot::Interned(idx) => {
            if let Some(Value::String(text)) = chunk.constants.get(*idx as usize).cloned() {
                declare_key_string(chunk, &text);
            }
            Some(*idx)
        }
        // Handled by the callers, which need the typeidx as well as the index.
        ResolvedSlot::Indexed { .. } => unreachable!("indexed slots bypass push_key"),
        ResolvedSlot::Dynamic(src) => {
            push_value(chunk, src, line);
            None
        }
    }
}

/// Push a value described by a [`ValueSource`]. Exposed so owner-adjacent
/// builders (`errors::emit_exception_new`) can arrange their own stack without
/// re-deriving the const dispatch.
pub fn push_value_public(chunk: &mut Chunk, val: &ValueSource, line: u32) {
    push_value(chunk, val, line);
}

/// Push the key as a runtime VALUE — for the host calls (`hasOwn`, `delete`)
/// that take the key as an operand rather than an immediate.
fn push_key_as_value(chunk: &mut Chunk, slot: &ResolvedSlot, line: u32) {
    match slot {
        // `hasOwn`/`delete` are host calls over the string-keyed property bag.
        // An indexed field is not in that bag at all — its presence is a static
        // fact about the type, decidable at compile time. Until that fold
        // exists, fall back to the declared name.
        ResolvedSlot::Indexed { .. } => {
            chunk.emit_string_const("", line);
        }
        ResolvedSlot::Key(name) => chunk.emit_string_const(name, line),
        // An interned key is a pool entry, not a pushable constant; the host
        // call needs the string itself, so it goes through the string-constant
        // import path like any other literal.
        ResolvedSlot::Interned(idx) => {
            let v = chunk.constants[*idx as usize].clone();
            if let Value::String(sname) = v {
                chunk.emit_string_const(&sname, line);
            }
        }
        ResolvedSlot::Dynamic(src) => push_value(chunk, src, line),
    }
}

/// Send a read result to its destination.
fn finish_dest(chunk: &mut Chunk, dest: Dest, line: u32) {
    match dest {
        Dest::Stack => {}
        Dest::Local(slot) => chunk.emit_op_u16(Op::LOCAL_SET, slot, line),
    }
}

/// Read a slot.
pub fn emit_class_get(
    chunk: &mut Chunk,
    obj: ObjSource,
    slot: &ResolvedSlot,
    dest: Dest,
    line: u32,
) {
    if let ResolvedSlot::Indexed { typeidx, field } = slot {
        push_obj(chunk, obj, line);
        chunk.emit_struct_field_op(Op::STRUCT_GET, *typeidx, *field, line);
        finish_dest(chunk, dest, line);
        return;
    }
    push_obj(chunk, obj, line);
    match push_key(chunk, slot, line) {
        Some(key) => chunk.emit_struct_field_op(Op::STRUCT_GET, 0, key, line),
        // Dynamic key: the key is a runtime value, so the read is a host call.
        None => host::emit(chunk, "ecma:object", "get", 2, line),
    }
    finish_dest(chunk, dest, line);
}

/// Write a slot.
///
/// `STRUCT_SET` pops `[obj, val]` and pushes nothing, so nothing is left behind
/// and no keep-or-drop decision reaches the caller. When the value is already
/// on the stack and the object is not, the object has to be moved underneath
/// it — the one case that needs a spill, and it happens HERE rather than at
/// every call site.
pub fn emit_class_set(
    chunk: &mut Chunk,
    obj: ObjSource,
    slot: &ResolvedSlot,
    val: ValueSource,
    line: u32,
) {
    // An INDEXED write is the same stack shape as a keyed one — the typeidx
    // and field index are both immediates — so only the operand arrangement is
    // shared, not the key push.
    if let ResolvedSlot::Indexed { typeidx, field } = slot {
        match (obj, &val) {
            (ObjSource::Stack, ValueSource::Stack) => {}
            (ObjSource::Local(o), ValueSource::Stack) => {
                let tmp = chunk.alloc_scratch(1);
                chunk.emit_op_u16(Op::LOCAL_SET, tmp, line);
                chunk.emit_op_u16(Op::LOCAL_GET, o, line);
                chunk.emit_op_u16(Op::LOCAL_GET, tmp, line);
            }
            (_, _) => {
                push_obj(chunk, obj, line);
                push_value(chunk, &val, line);
            }
        }
        chunk.emit_struct_field_op(Op::STRUCT_SET, *typeidx, *field, line);
        return;
    }
    match slot {
        // Returned above.
        ResolvedSlot::Indexed { .. } => unreachable!(),
        // A STATIC key is an IMMEDIATE on `STRUCT_SET`, not a stack operand, so
        // the stack only ever needs [obj, val].
        ResolvedSlot::Key(_) | ResolvedSlot::Interned(_) => {
            let key = push_key(chunk, slot, line).expect("static key");
            match (obj, &val) {
                // Already [obj, val]. Nothing to arrange.
                (ObjSource::Stack, ValueSource::Stack) => {}
                // [val] with the object in a local: the object has to go
                // UNDERNEATH the value, which is the one case needing a spill.
                //
                // ⚠ `alloc_scratch` shares an index space with the walker's
                // named locals and there is no scratch rewind, so this slot is
                // never reclaimed. Known allocator debt, audited at this ONE
                // site instead of at every caller a stack-only API would have
                // forced to hand-roll it.
                (ObjSource::Local(o), ValueSource::Stack) => {
                    let tmp = chunk.alloc_scratch(1);
                    chunk.emit_op_u16(Op::LOCAL_SET, tmp, line);
                    chunk.emit_op_u16(Op::LOCAL_GET, o, line);
                    chunk.emit_op_u16(Op::LOCAL_GET, tmp, line);
                }
                // The value is produced on demand, so it lands above the
                // object either way.
                (_, _) => {
                    push_obj(chunk, obj, line);
                    push_value(chunk, &val, line);
                }
            }
            chunk.emit_struct_field_op(Op::STRUCT_SET, 0, key, line);
        }
        // A DYNAMIC key is a real stack operand: the host call wants
        // [obj, key, val], so the key has to be inserted BETWEEN them.
        ResolvedSlot::Dynamic(key_src) => {
            match (obj, &val) {
                (ObjSource::Stack, ValueSource::Stack) => {
                    let tmp = chunk.alloc_scratch(1);
                    chunk.emit_op_u16(Op::LOCAL_SET, tmp, line);
                    push_value(chunk, key_src, line);
                    chunk.emit_op_u16(Op::LOCAL_GET, tmp, line);
                }
                (ObjSource::Local(o), ValueSource::Stack) => {
                    let tmp = chunk.alloc_scratch(1);
                    chunk.emit_op_u16(Op::LOCAL_SET, tmp, line);
                    chunk.emit_op_u16(Op::LOCAL_GET, o, line);
                    push_value(chunk, key_src, line);
                    chunk.emit_op_u16(Op::LOCAL_GET, tmp, line);
                }
                (_, _) => {
                    push_obj(chunk, obj, line);
                    push_value(chunk, key_src, line);
                    push_value(chunk, &val, line);
                }
            }
            host::emit(chunk, "ecma:object", "set", 3, line);
        }
    }
}

/// Emit the write for a key already arranged by the caller.
///
/// `emit_class_construct` owns its own stack discipline — it `dup`s the object
/// and pushes exactly [obj, (key,) val] per field — so it needs the raw write
/// rather than `emit_class_set`'s arrangement logic.
fn emit_set_op(chunk: &mut Chunk, key: Option<u16>, line: u32) {
    match key {
        Some(key) => chunk.emit_struct_field_op(Op::STRUCT_SET, 0, key, line),
        None => host::emit(chunk, "ecma:object", "set", 3, line),
    }
}

/// Construct an object with N named fields.
///
/// The ONLY site that builds a whole object, and therefore the only site where
/// field ORDER is decided — which is what makes it the M6 conversion point.
/// Today it emits `struct.new 0` plus N sets; after M6 it emits
/// `struct.new <typeidx>` with operands in the declared field order, and no
/// caller changes.
///
/// `ty` is a NAME rather than a `&ClassType` because the callers have a name:
/// the drawing value types are registered through a builder and declare no
/// fields at all, so there is no field ordering to pass yet. Resolving the name
/// here is what lets M6 supply one later without touching call sites.
pub fn emit_class_construct(
    chunk: &mut Chunk,
    ty: &str,
    fields: &[(ResolvedSlot, ValueSource)],
    line: u32,
) {
    // Values are taken into scratch first: `struct.new` pushes a fresh object
    // and any stack-sourced arguments are already above where it needs to go.
    let spills: Vec<Option<u16>> = fields
        .iter()
        .map(|(_, v)| match v {
            ValueSource::Stack => Some(chunk.alloc_scratch(1)),
            _ => None,
        })
        .collect();
    for slot in spills.iter().rev().flatten() {
        chunk.emit_op_u16(Op::LOCAL_SET, *slot, line);
    }

    chunk.emit_struct_new(0, 0, line);

    for ((slot, val), spill) in fields.iter().zip(&spills) {
        chunk.emit_dup(line);
        let key = push_key(chunk, slot, line);
        match spill {
            Some(tmp) => chunk.emit_op_u16(Op::LOCAL_GET, *tmp, line),
            None => push_value(chunk, val, line),
        }
        emit_set_op(chunk, key, line);
    }

    // The type stamp. `TypeIdentity` is a slot like any other, so it converts
    // with the rest rather than staying a bare string.
    chunk.emit_dup(line);
    let type_key = push_key(
        chunk,
        &ResolvedSlot::Key("__type".to_string()),
        line,
    );
    chunk.emit_string_const(ty, line);
    emit_set_op(chunk, type_key, line);
}

/// Allocate a bare object with no declared fields.
///
/// `struct.new 0` with a zero field count is the class model's ALLOCATOR — the
/// dynamic, string-keyed object form. 31 sites in `crates/` alone allocate this
/// way and then stamp fields individually, which is why they are not
/// `emit_class_construct` callers: the fields are not known at the allocation
/// point.
///
/// ▶▶ At M6 this becomes `struct.new_default <typeidx>` when the allocating
/// site knows its type, and stays typeidx 0 only for genuinely dynamic objects
/// (a JS object literal, a lua table). Routing it through the owner is what
/// makes that a change in one file rather than 31.
pub fn emit_class_alloc(chunk: &mut Chunk, line: u32) {
    chunk.emit_struct_new(0, 0, line);
}

/// Test whether a slot is present.
///
/// ⚠ There is no `STRUCT_HAS` in WASM GC — `opcode/gc.rs` defines NEW,
/// NEW_DEFAULT, GET, GET_S, GET_U, SET, NEW_DESC and NEW_DEFAULT_DESC and
/// nothing else — so this is inherently a host call, and these sites are NOT
/// struct emissions and never appear in the emission count.
///
/// At M6, `has` against a DECLARED struct type is decidable at compile time:
/// the field is in the type or it is not. This should collapse to a constant
/// there rather than convert.
/// ⚠ **Correct only for a KEYED slot.** `push_key_as_value` pushes the slot's
/// STRING key, so for `ResolvedSlot::Indexed` this asks a string-keyed property
/// probe about a field that lives in an indexed struct slot — and indexed
/// storage never populates the property map, so the answer is a confident
/// `false` for a field that is demonstrably present.
///
/// That exact split — a guard resolving differently from the read it guards —
/// was 78 of the js regressions via `emit_js_private_brand_check`, and the fix
/// there was `ref.test`, which is what "does this object have this class's
/// field" actually means. The single caller today
/// (`python/repr_adapter.rs`) passes a keyed slot, so this is sound as used;
/// it is documented rather than fixed because the correct answer depends on the
/// slot kind and there is no second caller to generalise from yet.
pub fn emit_class_has(
    chunk: &mut Chunk,
    obj: ObjSource,
    slot: &ResolvedSlot,
    dest: Dest,
    line: u32,
) {
    push_obj(chunk, obj, line);
    push_key_as_value(chunk, slot, line);
    host::emit(chunk, "ecma:object", "hasOwn", 2, line);
    finish_dest(chunk, dest, line);
}


// ── The compiler-side facade ────────────────────────────────────────────
//
// The free functions above take a `&mut Chunk` because that is the shape the
// ~180 emitter-side wrappers are called in. Inside the compiler the receiver is
// `&mut Compiler`, and threading `&mut self.chunks[self.current]`, `self.line`
// and a `SlotNames` through every call site would be the same hand-rolled
// plumbing this API exists to delete. These four methods are the compiler's
// spelling of the same five verbs.

impl crate::Compiler {
    /// Read a slot, leaving the value on the stack.
    pub(crate) fn class_get(&mut self, obj: ObjSource, slot: &ClassSlot) {
        let line = self.line;
        if self.emit_guarded_indexed_get(obj, slot, Dest::Stack, line) {
            return;
        }
        let slot = self.resolve_slot(slot);
        emit_class_get(&mut self.chunks[self.current], obj, &slot, Dest::Stack, line);
    }

    /// Read a slot into a local.
    pub(crate) fn class_get_to(&mut self, obj: ObjSource, slot: &ClassSlot, dest: u16) {
        let line = self.line;
        if self.emit_guarded_indexed_get(obj, slot, Dest::Local(dest), line) {
            return;
        }
        let slot = self.resolve_slot(slot);
        emit_class_get(
            &mut self.chunks[self.current],
            obj,
            &slot,
            Dest::Local(dest),
            line,
        );
    }

    /// An indexed READ, guarded by `ref.test` against the declaring class.
    ///
    /// ⛔ A DECLARED TYPE IS NOT A GUARANTEE ABOUT THE RUNTIME OBJECT, and an
    /// unguarded `struct.get <typeidx>` on something that is not that type
    /// TRAPS rather than answering `undefined`. `seam3_indexable` rules out
    /// subclassing (`!has_parent`) and nothing else — not a `with`-expression
    /// copy built by a helper, not a platform value, not `null`. Measured: 11
    /// csharp tests, `trap: struct.get field index out of range` and
    /// `trap: struct.get on a non-struct`.
    ///
    /// The WRITE side needs no guard: it runs inside the constructor, where the
    /// receiver is `this` and therefore exact and non-null.
    ///
    /// `ref.test` is the spec instruction for exactly this question, so the
    /// fast path stays a real `struct.get` and the miss falls back to the
    /// string key the read would have used anyway.
    ///
    /// Returns whether it emitted; `false` means the caller takes the ordinary
    /// path.
    fn emit_guarded_indexed_get(
        &mut self,
        obj: ObjSource,
        slot: &ClassSlot,
        dest: Dest,
        line: u32,
    ) -> bool {
        let ClassSlot::InstanceField { class: Some(class), field } = slot else {
            return false;
        };
        let ResolvedSlot::Indexed { typeidx, field: fieldidx } = self.resolve_slot(slot) else {
            return false;
        };
        let class = class.clone();
        // The key this read would have used without indexing — the SAME answer
        // `storage_key` gives, so the miss path is byte-for-byte the old one.
        let fallback = self.resolve_slot_interned(&ClassSlot::internal(
            self.js_member_storage_name_for_class(&class, field),
        ));
        // A named local, not `alloc_scratch`: the receiver is read twice (once
        // to test, once to load) and scratch shares an index space with the
        // walker's named locals.
        let recv = self.define_local("__seam3_recv");
        push_obj(&mut self.chunks[self.current], obj, line);
        self.emit_u16(Op::LOCAL_SET, recv);

        self.emit_u16(Op::LOCAL_GET, recv);
        self.emit_ref_type_test(Op::REF_TEST, &class, line);
        self.chunk().emit_if_value(line);
        self.emit_u16(Op::LOCAL_GET, recv);
        self.chunk()
            .emit_struct_field_op(Op::STRUCT_GET, typeidx, fieldidx, line);
        self.chunk().emit_else(line);
        self.emit_u16(Op::LOCAL_GET, recv);
        emit_class_get(
            &mut self.chunks[self.current],
            ObjSource::Stack,
            &fallback,
            Dest::Stack,
            line,
        );
        self.chunk().emit_end(line);

        finish_dest(&mut self.chunks[self.current], dest, line);
        true
    }

    /// Write a slot.
    pub(crate) fn class_set(&mut self, obj: ObjSource, slot: &ClassSlot, val: ValueSource) {
        let line = self.line;
        let slot = self.resolve_slot(slot);
        emit_class_set(&mut self.chunks[self.current], obj, &slot, val, line);
    }

    /// The name resolver for slots emitted from compiler context.
    ///
    /// Private mangling depends on compiler state, so a slot that names a
    /// DECLARED member resolves through the compiler; engine-internal keys and
    /// the fixed spellings (`prototype`, `__proto__`, `__type`) never mangle
    /// and resolve plainly. `storage_key` routes on the variant, so this only
    /// has to answer for the member cases.
    /// Resolve and intern in one step, for a key the original code hoisted
    /// above its emit sites. See [`resolve_interned`] — `add_constant` does not
    /// de-duplicate, so a hoisted key must intern exactly once.
    /// Allocate a bare object. See [`emit_class_alloc`].
    pub(crate) fn class_alloc(&mut self) {
        let line = self.line;
        emit_class_alloc(&mut self.chunks[self.current], line);
    }

    pub(crate) fn resolve_slot_interned(&mut self, slot: &ClassSlot) -> ResolvedSlot {
        let resolved = self.resolve_slot(slot);
        match resolved {
            ResolvedSlot::Key(name) => {
                let idx = self.chunks[self.current]
                    .add_constant(Value::String(Arc::from(name.as_str())));
                ResolvedSlot::Interned(idx)
            }
            other => other,
        }
    }

    /// Read a slot that is already resolved.
    pub(crate) fn class_get_resolved(&mut self, obj: ObjSource, slot: &ResolvedSlot) {
        let line = self.line;
        emit_class_get(&mut self.chunks[self.current], obj, slot, Dest::Stack, line);
    }

    /// Write a slot that is already resolved.
    pub(crate) fn class_set_resolved(
        &mut self,
        obj: ObjSource,
        slot: &ResolvedSlot,
        val: ValueSource,
    ) {
        let line = self.line;
        emit_class_set(&mut self.chunks[self.current], obj, slot, val, line);
    }

    /// The declaring class of a member ACCESS, for seam-3 indexing.
    ///
    /// A write site knows its class from `current_class`; a read site has only
    /// the receiver expression, so the class comes from inference.
    /// `resolve_receiver_type_hint` is the superset of `infer_expr_type_hint`
    /// and is what every other receiver-typed path already uses.
    ///
    /// ⚠ Returns `None` freely. A receiver whose type cannot be inferred keeps
    /// the string key, which is always correct — indexed access is the
    /// optimisation, the property bag is the semantics.
    /// The registered type whose INDEXED field list contains `storage`, if
    /// exactly one does.
    ///
    /// This exists so a **guard can resolve the way its lookup resolves**. A
    /// presence check written as a string-keyed property probe answers `false`
    /// for a field that lives in an indexed struct slot, because indexed
    /// storage never populates the property map — the read succeeds and the
    /// guard in front of it throws.
    ///
    /// Asked as a question about the TABLE, not about a language: "is this
    /// storage name an indexed field of some class this compiler authored?"
    /// Nothing here inspects which language produced the class.
    ///
    /// Ambiguity is answered `None` rather than by picking: two classes with
    /// the same storage name give no single type to test against, and a wrong
    /// type test is worse than falling back to the string probe.
    /// The class this compiler DECLARED as holding `storage`, whether or not
    /// that class also holds it as an indexed field.
    ///
    /// ⛔ THE LICENCE IS THE WRONG GATE FOR A TYPE TEST.
    /// [`Self::indexed_owner_of_storage`] answers a storage question — "may I
    /// emit `struct.get` against this field index" — and `seam3_indexable` is
    /// exactly the right filter for that. A js private BRAND asks something
    /// else entirely: *was this object constructed by this class*, which is
    /// `ref.test`, and `ref.test` cares only that the type exists. Whether the
    /// field went to an indexed slot or to the property map has no bearing on
    /// it.
    ///
    /// The two were the same function, so withholding the indexing licence
    /// silently withdrew the brand's type test too: `#tokenVal in proxy`
    /// dropped to a property probe, which forwards through a Proxy to its
    /// target and answered `true` where ECMA-262 §13.10.1 requires `false` — a
    /// Proxy carries no private fields of its own and does not inherit the
    /// target's.
    ///
    /// Ambiguity still yields `None`: two authored types holding one storage
    /// name cannot pick a type to test against.
    pub(crate) fn declaring_owner_of_storage(&self, storage: &str) -> Option<String> {
        let mut found: Option<&str> = None;
        for entry in &self.chunks[0].types {
            // ⛔ NOT unfiltered. A `ref.test` is only an answer about identity
            // where the class's own typeidx is what the allocation used —
            // `rtt_testable`. Scanning every published type instead made
            // `#tag in s` answer `false` for a genuine `Sub` instance, because
            // `new Sub()` allocates with `Base`'s rtt and the test was emitted
            // against `Sub`. Shadowed private names in a subclass are the shape
            // that catches it.
            if !self.rtt_testable.contains(&self.canon(&entry.name)) {
                continue;
            }
            if entry.fields.iter().any(|f| f == storage) {
                if found.is_some() {
                    return None;
                }
                found = Some(&entry.name);
            }
        }
        found.map(str::to_string)
    }

    pub(crate) fn indexed_owner_of_storage(&self, storage: &str) -> Option<String> {
        let mut found: Option<&str> = None;
        for entry in &self.chunks[0].types {
            if !self.seam3_indexable.contains(&self.canon(&entry.name)) {
                continue;
            }
            if entry.fields.iter().any(|f| f == storage) {
                if found.is_some() {
                    return None;
                }
                found = Some(&entry.name);
            }
        }
        found.map(str::to_string)
    }


    pub(crate) fn resolve_slot(&self, slot: &ClassSlot) -> ResolvedSlot {
        // ▶▶ SEAM 3.
        //
        // ⛔ WRITES AND READS MUST CONVERT TOGETHER. Turning this on for
        // declared fields emitted `struct.set <typeidx> <fieldidx>` into
        // `obj.fields` while every READ still went through the string-keyed
        // path into `obj.properties`. **There is no name→index bridge on the
        // read**: `dispatch.rs:4285` dual-writes properties→fields on a
        // string-keyed SET, and nothing does the reverse. `Object::get` never
        // consults `fields`.
        //
        // Measured, by name: **dart 693→772 (+79), js 782→900 (+118)** — every
        // one a field/mixin/factory/covariant test, i.e. exactly the classes
        // that got a real typeidx. Four hand-written probes passed first, which
        // is why this is gated on a corpus and not on a repro.
        //
        // A slot arrives here naming its declaring class from exactly two
        // places: the field initializer (`classes.rs`) on the write, and
        // `member_slot_for_receiver` (`metadata.rs`) on the read. Every OTHER
        // member site pre-resolves its key to a string and arrives as
        // `Internal`, which has no class to index against and therefore stays
        // on the string path by construction — that is the shape to look for
        // when a read the type clearly declares is still string-keyed.
        // `indexed_field` resolves BY NAME and returns `None` for anything the
        // type does not declare, which is what keeps dynamic objects there.
        if let ClassSlot::InstanceField { class: Some(class), field } = slot {
            if let Some(indexed) = self.indexed_field(class, field) {
                return indexed;
            }
        }
        resolve(slot, &CompilerNames(self))
    }

    /// Resolve a declared field to `(typeidx, fieldidx)` — **by name**.
    ///
    /// ⚠ THE LOOKUP IS BY NAME AND MUST STAY THAT WAY. A type's declared field
    /// list can carry entries the language synthesised: python's `@dataclass`
    /// emits `0:__class__ 1:x 2:y 3:label 4:__dataclass_fields__
    /// 5:__match_args__`, so `x` sits at index **1**. Indexing by declaration
    /// POSITION would write it into `__class__`, silently, because nothing
    /// validates the slot afterwards. Length checks fail the same way from the
    /// other side — 6 != 3 would reject a perfectly indexable class.
    ///
    /// ⚠ Allocation and access must agree. `struct.get <typeidx> <idx>` reads
    /// `obj.fields[idx]` and TRAPS on a short vec, so this may only return
    /// `Indexed` for a type whose instances are allocated with the same
    /// typeidx — which is exactly the condition `classes.rs`'s
    /// `if typeidx != 0 { STRUCT_NEW_DEFAULT }` already branches on. A type
    /// reserved but never populated keeps an empty `fields`, finds nothing
    /// here, and stays on the string path.
    fn indexed_field(&self, class: &str, field: &str) -> Option<ResolvedSlot> {
        let storage = self.js_member_storage_name_for_class(class, field);
        let canon_class = self.canon(class);
        // ⛔ The licence check. See `Compiler::seam3_indexable`.
        if !self.seam3_indexable.contains(&canon_class) {
            return None;
        }
        let pos = self.chunks[0]
            .types
            .iter()
            .position(|t| t.name == canon_class)?;
        let entry = &self.chunks[0].types[pos];
        let idx = entry.fields.iter().position(|f| *f == storage)?;
        Some(ResolvedSlot::Indexed {
            // `reserve_type_slot` hands out 1-based indices; typeidx 0 is the
            // dynamic form, so the +1 is the same convention, not an offset.
            typeidx: pos as u16 + 1,
            field: idx as u16,
        })
    }
}

/// Borrowed view of the compiler for name resolution, so `class_set` can hold
/// `&mut self.chunks` and a `&self` resolver at the same time.
struct CompilerNames<'a>(&'a crate::Compiler);

impl SlotNames for CompilerNames<'_> {
    /// ⛔ DELEGATES. This re-implemented `impl SlotNames for Compiler`
    /// verbatim — the same two-arm match, differing only by a `.0` — so the
    /// private-mangling rule had two implementations that had to be kept in
    /// step by hand. The wrapper exists for a BORROW reason (emission needs
    /// `&mut self.chunks` while resolution needs `&self`), which is a reason to
    /// wrap, never a reason to restate the rule.
    fn storage_name(&self, class: Option<&str>, field: &str) -> String {
        self.0.storage_name(class, field)
    }
}

use crate::opcode::Op;
use crate::value::Value;
use std::collections::BTreeMap;

/// The namespace designated for imported string constants
/// (js-string-builtins, § String constants). Every import from it is a global
/// of type `(ref extern)` whose value **is** the import's field name.
pub const STRING_CONSTANTS_MODULE: &str = "wasm:string-constants";

/// How an imported global is spelled in the VM's name-keyed global map.
///
/// WASM addresses globals by index and imported globals are identified by
/// their `(module, name)` pair; our map is keyed by name, and that name space
/// is shared with user variables. Spelling the pair is what keeps the two
/// disjoint: `var count = 5` must not be able to redefine the string constant
/// `"count"`, and no source language can declare an identifier containing
/// `:`.
pub fn imported_global_key(module: &str, name: &str) -> String {
    format!("{}::{}", module, name)
}

/// The namespace a module declares its HOST-provided global bindings under.
///
/// A free global — a name the module reads but never writes — is an import,
/// and this is the module it comes from: the embedder. `env` rather than a
/// Vybe-specific name because the codebase already recognises it as one of the
/// embedder-provided namespaces (`bundle.rs` groups `"*"`, `"env"` and
/// `"wasm:string-constants"`), and because it is the de-facto WASM convention
/// a JS host already supplies without being taught anything.
pub const HOST_GLOBALS_MODULE: &str = "env";

/// A host function import declaration — (module, name).
/// Like WASM: (import "vybe:math" "floor" (func ...))
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct Import {
    pub module: String,
    pub name: String }

/// A spec EH tag declaration (exception-handling proposal tag section):
/// `(tag $t (param ...))`. Identity is the DECLARATION — two entries are
/// distinct tags even with equal names/arities. `debug_name` only feeds
/// diagnostics and import/export naming; `arity` is the payload count of
/// the tag's function-type signature (result type must be empty per spec).
#[derive(Debug, Clone)]
pub struct TagDecl {
    pub debug_name: String,
    pub arity: u8,
    /// `true` for tag IMPORTS: per spec, imports resolve to an existing
    /// tag entity by name at instantiation (same name ⇒ same entity —
    /// this is how `vybe:exception` is shared across chunks/modules).
    /// `false` for local declarations, which are always fresh entities.
    pub imported: bool }

/// Property descriptor per ECMA-262 §6.2.4 — represented using WASM Annotations proposal.
/// Format: (@ecma262 descriptor field_name writable enumerable configurable)
/// Serialized as flags byte in custom section: bit 0=writable, bit 1=enumerable, bit 2=configurable
/// All three flags default to true (most permissive) if not specified.
/// Fully compliant with proposals/annotations/proposals/annotations/Overview.md
#[derive(Debug, Clone, PartialEq, Eq)]
pub struct PropertyDescriptor {
    /// Can the property value be changed (§7.3.2 Put, SetPropertyValue).
    pub writable: bool,
    /// Can the property appear in for-in loops and Object.keys / Object.getOwnPropertyNames.
    pub enumerable: bool,
    /// Can the property be deleted (§7.3.7 DeletePropertyOrThrow).
    pub configurable: bool }

impl PropertyDescriptor {
    /// All attributes true (standard object property).
    pub fn standard() -> Self {
        PropertyDescriptor {
            writable: true,
            enumerable: true,
            configurable: true }
    }

    /// Read-only non-enumerable non-configurable (built-in method/property).
    pub fn builtin() -> Self {
        PropertyDescriptor {
            writable: false,
            enumerable: false,
            configurable: false }
    }

    /// Non-enumerable but writable and configurable (error.message).
    pub fn non_enumerable() -> Self {
        PropertyDescriptor {
            writable: true,
            enumerable: false,
            configurable: true }
    }
}

/// The WASM GC composite-type shape of a defined type (spec: `comptype`).
/// A defined type is exactly one of these. This is the runtime rtt discriminant
/// `ref.test`/`ref.cast` and `array.*`/`struct.*` branch on — an array ref traps
/// on null/out-of-bounds, a struct ref carries named fields. Defaults to `Struct`
/// so every existing class/interface registration (which never sets it) is
/// unaffected.
#[derive(Debug, Clone, Copy, PartialEq, Eq, Default)]
pub enum CompositeKind {
    /// `(struct …)` — named fields; also the default for class/interface types.
    #[default]
    Struct,
    /// `(array …)` — homogeneous indexed storage; `array.get`/`set`/`copy` trap
    /// on null ref or out-of-bounds index (WASM GC proposal §array).
    Array,
    /// `(func …)` — a function signature type.
    Func }

/// A compile-time type definition — WASM GC type section entry.
/// Describes a class/struct with named fields and vtable methods.
/// Loaded into TypeRegistry before execution.
#[derive(Debug, Clone)]
pub struct TypeEntry {
    pub name: String,
    /// WASM GC composite shape (struct / array / func). Determines the runtime
    /// rtt kind stamped onto instances so `array.*` ops trap per spec.
    pub kind: CompositeKind,
    /// Declared supertype — `sub $i`, a **1-based index into this same type
    /// table**; `0` is `sub` with an empty typeidx list, i.e. no declared
    /// supertype.
    ///
    /// This used to be a NAME resolved at load. An index is what the spec
    /// declares and it removes a bug class outright: a parent spelled slightly
    /// differently from its declaration used to resolve to nothing and the
    /// class landed with no supertype at all. A supertype defined elsewhere
    /// (a host builtin, another component) still gets an entry here — a
    /// declaration bound at load — so the reference is an index either way.
    pub parent_index: u16,
    /// Field names in order. Field i is at `fields[i]` in the object's indexed storage.
    pub fields: Vec<String>,
    /// Vtable: method_name → chunk_index. Methods are shared across all instances.
    pub methods: Vec<(String, usize)>,
    /// Whether this is an interface definition (not a concrete class).
    pub is_interface: bool,
    /// Interfaces this type implements — **1-based indices into this same
    /// type table**, like `parent_index`. An interface defined elsewhere gets
    /// a declaration here and binds at load, so the link is an index either
    /// way and nothing resolves an interface by name at run time.
    pub implements: Vec<u16>,
    /// Constructor chunk index (if any). Resolved during load_type_table.
    pub constructor_chunk: Option<usize>,
    /// Field property descriptors (WASM Annotations proposal @ecma262 namespace).
    /// Maps field_name → descriptor. Fields without entries default to PropertyDescriptor::standard().
    pub field_descriptors: std::collections::HashMap<String, PropertyDescriptor> }

/// A constant initialization expression (Extended Const Expressions proposal).
/// Evaluated at module instantiation time, before code execution.
#[derive(Debug, Clone)]
pub enum ConstExpr {
    /// A literal value.
    Value(Value),
    /// Reference another global: global_name.
    GlobalGet(String),
    /// Add two const exprs: left + right.
    Add(Box<ConstExpr>, Box<ConstExpr>),
    /// Multiply two const exprs: left * right.
    Mul(Box<ConstExpr>, Box<ConstExpr>),
    /// Create a function reference from a chunk index.
    /// Evaluated at load time — produces a callable Function object.
    RefFunc(usize) }

/// A global variable initializer — evaluated at link/load time.
#[derive(Debug, Clone)]
pub struct GlobalInit {
    /// Global name (stored in VM.globals).
    pub name: String,
    /// Initialization expression (evaluated before code runs).
    pub init: ConstExpr }

/// A continuation tag — defines the type contract for typed continuations.
#[derive(Debug, Clone)]
pub struct ContinuationTag {
    /// Tag name (for debugging and cross-language matching).
    pub name: String,
    /// Expected yield value type name (empty = any).
    pub yield_type: String,
    /// Expected resume value type name (empty = any).
    pub resume_type: String }

#[derive(Debug, Clone)]
pub struct ActiveDataSegment {
    pub memory_index: u32,
    pub offset: u64,
    pub data_index: u32 }

#[derive(Debug, Clone)]
pub struct ActiveElementSegment {
    pub table_index: u32,
    pub offset: u64,
    pub elem_index: u32 }

#[derive(Debug, Clone, PartialEq, Eq)]
pub struct StackSwitchHandler {
    /// 0 = on-tag-to-label, 1 = on-tag-to-switch.
    pub kind: u8,
    pub tag_index: u32,
    /// For decoded standard Wasm this is the structural label index.
    /// Direct bytecode VM tests use a resolved bytecode instruction offset.
    pub label_index: u32 }

/// A compiled chunk of bytecode — one per function/script.
#[derive(Debug, Clone)]
pub struct Chunk {
    pub code: Vec<u8>,
    pub constants: Vec<Value>,
    /// `string → position in `constants`` for [`Chunk::intern_string_constant`].
    ///
    /// The pool is APPEND-ONLY — `add_constant` only pushes, and nothing
    /// removes, reorders or rewrites it — so a position recorded here stays
    /// valid for the life of the chunk. Without this the interner scanned the
    /// whole pool on every call, and since every global access and every
    /// string constant now routes through it, that was quadratic in the
    /// number of constants.
    string_constants: std::collections::HashMap<String, u16>,
    pub lines: Vec<u32>,
    pub name: String,
    pub arity: u8,
    pub local_count: u16,
    /// Debug metadata: `(slot, name)` for every local variable the compiler
    /// defined in this chunk (params + block locals). Slots may repeat when
    /// sibling blocks reuse a slot; consumers (the debugger) take the last
    /// name per slot. Empty when the frontend emitted no debug names. Never
    /// read during execution — inspection/eval only.
    pub local_names: Vec<(u16, String)>,
    /// Debug metadata (script chunk only): bytecode offset where the user's
    /// own code begins, i.e. just past the injected language runtime prelude
    /// (`__vybe_user_code_start__` marker). `None` when there is no prelude or
    /// the frontend didn't mark it. The step debugger uses it to land the first
    /// pause in user code and to skip/step-over "system" (prelude) code. Never
    /// read during execution.
    pub user_code_offset: Option<u32>,
    /// Import table — only on the script chunk (chunk 0).
    /// Each entry is a (module, name) pair.
    /// CallHost operand indexes into this table.
    pub imports: Vec<Import>,
    /// Imported **globals** — a separate index space from `imports`, exactly
    /// as in WASM, where each import kind numbers independently and only
    /// function imports are reachable by `call`. Keeping them apart is what
    /// stops a declared global from shifting the function indices that
    /// `CALL_IMPORT` operands carry.
    pub global_imports: Vec<Import>,
    /// Type table — WASM GC type section. Only on the script chunk (chunk 0).
    /// Each entry defines a class type with fields and vtable methods.
    /// Loaded into VM's TypeRegistry before execution.
    pub types: Vec<TypeEntry>,
    /// Exception tag table — maps tag index to type name for typed exception handling.
    /// Tag 0 = catch-all (matches any exception).
    /// Tag N = matches exceptions whose type is `exception_tags[N]` or a subtype.
    /// LEGACY: superseded by `tags` (spec EH tag section); removed once the
    /// name-matching catch path is gone.
    pub exception_tags: Vec<String>,
    /// Spec EH tag section (exception-handling proposal): each entry is a
    /// FRESH tag entity — identity is the declaration, `debug_name` is
    /// metadata only (import/export naming), never used for matching.
    /// `arity` is the payload value count (the tag's signature params).
    pub tags: Vec<TagDecl>,
    /// Type imports — types from other components this chunk needs.
    /// Each entry is (interface_name, type_name).
    pub type_imports: Vec<(String, String)>,
    /// Type exports — types this component makes available to others.
    /// Each entry is (interface_name, type_name, type_id).
    pub type_exports: Vec<(String, String, usize)>,
    /// Global initializers — const expressions evaluated at load time.
    pub global_inits: Vec<GlobalInit>,
    /// Continuation tags — typed contracts for suspend/resume.
    pub continuation_tags: Vec<ContinuationTag>,
    /// Declared linear memories for modules decoded from standard WASM.
    /// Entries are minimum page counts, in spec memory-index order.
    pub memory_min_pages: Vec<u64>,
    /// Declared linear memory maximums for modules decoded from standard WASM.
    /// Entries align with `memory_min_pages`; `None` means unbounded by the
    /// module type.
    pub memory_max_pages: Vec<Option<u64>>,
    /// Per-memory index type: `true` = 64-bit (memory64), `false` = 32-bit.
    /// Aligns with `memory_min_pages`. Memory64 adds NO new opcodes — the VM
    /// reads this at each load/store to decide the address width (i32 vs i64)
    /// and memarg-offset width, exactly as the spec (`C.mems[i] = at limits`).
    pub memory_is_64: Vec<bool>,
    /// Declared reference tables for modules decoded from standard WASM.
    /// Entries are minimum element counts, in spec table-index order.
    pub table_min_sizes: Vec<u64>,
    /// Optional maximum element count per table (aligns with `table_min_sizes`).
    /// `table.grow` past this returns -1 (WASM spec). `None` = unbounded.
    pub table_max_sizes: Vec<Option<u64>>,
    /// Per-table index type: `true` = 64-bit (table64). Aligns with
    /// `table_min_sizes`. Like memory64, table64 adds no new opcodes — the
    /// VM reads this to pick the index operand width (i32 vs i64).
    pub table_is_64: Vec<bool>,
    /// Passive data segment payloads decoded from standard WASM.
    pub data_segments: Vec<Vec<u8>>,
    /// Passive element segment payloads decoded from standard WASM.
    pub elem_segments: Vec<Vec<Value>>,
    /// Passive element segments compiled from source (`(elem $e $f …)`), stored
    /// as per-segment lists of function chunk indices. The VM resolves each to a
    /// funcref `Value` at instantiation (same as `REF_FUNC`) and populates
    /// `elem_segments`, so `table.init`/`array.new_elem` read real funcrefs.
    pub passive_elem_funcs: Vec<Vec<usize>>,
    /// Active data segments to instantiate before executing decoded WASM.
    pub active_data_segments: Vec<ActiveDataSegment>,
    /// Active element segments to instantiate before executing decoded WASM.
    pub active_elem_segments: Vec<ActiveElementSegment>,
    /// Stack-switching handler vectors keyed by bytecode opcode offset.
    pub stack_switch_handlers: BTreeMap<usize, Vec<StackSwitchHandler>>,
    /// Number of results this function returns. Default 1 (single
    /// externref) matches the pre-multi-value ABI. A chunk that wants
    /// to take advantage of the multi-value proposal sets this >1; the
    /// WASM emitter then declares the function type with that many
    /// externref results.
    pub result_arity: u8,
    /// True when this chunk is a compiled method — its `arity` includes an
    /// implicit leading receiver slot (`self`/`this`). Imported WASM functions
    /// (no receiver) leave this false.
    pub is_method: bool,
    /// The function's declared WASM parameter count (user params only, no
    /// receiver). Paired with `result_arity`, this is the function's type
    /// "shape" — the `call_indirect` runtime type check compares it against
    /// the call site's expected `[params]→[results]`.
    pub param_count: u8,
    /// JSPI: this function is `async` in its source language. The
    /// compiler sets this flag when compiling an `async function` /
    /// `async def` / `Async Function`. The WASM emitter writes a
    /// `vybe.jspi` custom section listing every async chunk so a JS
    /// host can wrap the corresponding export with
    /// `WebAssembly.promising(fn)` and therefore have it return a real
    /// JS Promise that resolves when the Vybe fiber completes. Non-Vybe
    /// WASM engines ignore unknown custom sections, so setting this on
    /// a sync function is at worst no-op on those engines. On Vybe VM
    /// it's informational only — the runtime already knows via
    /// `PROMISE_SUSPEND` opcodes in the body.
    pub is_async: bool,
    /// When true, `call_value` on a function bound to this chunk
    /// returns a fresh generator continuation instead of executing the
    /// body. Compilers set this for `def` with `yield`, `function*`,
    /// C# `yield return`, Ruby `Enumerator::new`, Dart `sync*`. Lowers
    /// to WASM stack-switching `cont.new` at emit time.
    pub is_generator: bool,
    /// Shared scratch local for emit_dup (local.tee + local.get).
    pub dup_slot: Option<u16>,
    /// High-water mark for scratch locals allocated by emitter helpers.
    /// The compiler MUST set `local_count = local_count.max(scratch_high_water)`
    /// at function finalization. This ensures `alloc_scratch` allocations
    /// are always included regardless of which compilation path runs.
    pub scratch_high_water: u16,
    /// Number of captured variables (closures). The VM populates
    /// local slots [capture_base..capture_base+capture_count) from
    /// the function object's `__capture_N` properties at call time.
    pub capture_count: u8,
    /// First local slot for captured variables.
    pub capture_base: u16 }

impl Chunk {
    pub fn new(name: impl Into<String>) -> Self {
        Chunk {
            code: Vec::new(),
            constants: Vec::new(),
            string_constants: std::collections::HashMap::new(),
            lines: Vec::new(),
            name: name.into(),
            arity: 0,
            local_count: 0,
            local_names: Vec::new(),
            user_code_offset: None,
            dup_slot: None,
            imports: Vec::new(),
            global_imports: Vec::new(),
            types: Vec::new(),
            exception_tags: Vec::new(),
            tags: Vec::new(),
            type_imports: Vec::new(),
            type_exports: Vec::new(),
            global_inits: Vec::new(),
            continuation_tags: Vec::new(),
            memory_min_pages: Vec::new(),
            memory_max_pages: Vec::new(),
            memory_is_64: Vec::new(),
            table_min_sizes: Vec::new(),
            table_max_sizes: Vec::new(),
            table_is_64: Vec::new(),
            data_segments: Vec::new(),
            elem_segments: Vec::new(),
            passive_elem_funcs: Vec::new(),
            active_data_segments: Vec::new(),
            active_elem_segments: Vec::new(),
            stack_switch_handlers: BTreeMap::new(),
            result_arity: 1,
            is_method: false,
            param_count: 0,
            is_async: false,
            is_generator: false,
            capture_count: 0,
            capture_base: 0,
            scratch_high_water: 0 }
    }

    /// Add an import and return its index (used by CallHost operand).
    pub fn add_import(&mut self, module: impl Into<String>, name: impl Into<String>) -> u16 {
        let import = Import {
            module: module.into(),
            name: name.into() };
        // Deduplicate — return existing index if already imported
        for (i, existing) in self.imports.iter().enumerate() {
            if *existing == import {
                return i as u16;
            }
        }
        self.imports.push(import);
        (self.imports.len() - 1) as u16
    }

    /// Declare an imported **global** and return its index in the global
    /// import space. Deduplicated on `(module, name)` — two references to the
    /// same import are the same global, as in WASM.
    pub fn add_global_import(&mut self, module: impl Into<String>, name: impl Into<String>) -> u16 {
        let import = Import { module: module.into(), name: name.into() };
        for (i, existing) in self.global_imports.iter().enumerate() {
            if *existing == import {
                return i as u16;
            }
        }
        self.global_imports.push(import);
        (self.global_imports.len() - 1) as u16
    }

    /// Intern a string in the constant pool. The pool is read-only at run
    /// time, so sharing one entry between sites is always safe; string
    /// constants reference theirs once per emit site and there are ~1,455 of
    /// them in a large module.
    pub fn intern_string_constant(&mut self, s: &str) -> u16 {
        if let Some(&i) = self.string_constants.get(s) {
            return i;
        }
        let idx = self.add_constant(Value::String(std::sync::Arc::from(s)));
        self.string_constants.insert(s.to_string(), idx);
        idx
    }

    /// Declare a spec EH tag (exception-handling proposal tag section) and
    /// return its index. Every declaration is a FRESH entity — no
    /// deduplication: per spec, tags are "created fresh" and identity is
    /// the declaration itself, not the name. `debug_name` is metadata for
    /// diagnostics and import/export naming only.
    pub fn declare_exception_tag(&mut self, debug_name: impl Into<String>, arity: u8) -> u16 {
        self.tags.push(TagDecl {
            debug_name: debug_name.into(),
            arity,
            imported: false });
        (self.tags.len() - 1) as u16
    }

    /// Import a tag by name (spec tag import): resolves to the SAME entity
    /// as every other import of that name at load time. Deduplicated within
    /// the chunk — importing the same name twice is the same tag either way.
    pub fn import_exception_tag(&mut self, name: impl Into<String>, arity: u8) -> u16 {
        let name = name.into();
        for (i, t) in self.tags.iter().enumerate() {
            if t.imported && t.debug_name == name {
                return i as u16;
            }
        }
        self.tags.push(TagDecl {
            debug_name: name,
            arity,
            imported: true });
        (self.tags.len() - 1) as u16
    }

    /// Add an exception tag and return its index.
    /// Tag 0 = catch-all. Tag N = typed catch for exceptions matching `type_name`.
    /// LEGACY name-matching table — see `declare_exception_tag`.
    pub fn add_exception_tag(&mut self, type_name: impl Into<String>) -> u8 {
        let name = type_name.into();
        // Deduplicate
        for (i, existing) in self.exception_tags.iter().enumerate() {
            if *existing == name {
                return i as u8;
            }
        }
        self.exception_tags.push(name);
        (self.exception_tags.len() - 1) as u8
    }

    /// Add a global initializer (evaluated at load time).
    pub fn add_global_init(&mut self, name: impl Into<String>, init: ConstExpr) {
        self.global_inits.push(GlobalInit {
            name: name.into(),
            init });
    }

    /// Add a continuation tag and return its index.
    pub fn add_continuation_tag(
        &mut self,
        name: impl Into<String>,
        yield_type: impl Into<String>,
        resume_type: impl Into<String>,
    ) -> u16 {
        let tag = ContinuationTag {
            name: name.into(),
            yield_type: yield_type.into(),
            resume_type: resume_type.into() };
        self.continuation_tags.push(tag);
        (self.continuation_tags.len() - 1) as u16
    }

    pub fn emit(&mut self, byte: u8, line: u32) {
        self.code.push(byte);
        self.lines.push(line);
    }

    /// Emit an opcode's 4 bytes. Ops WITH immediates legitimately come through
    /// here too — `emit_i32_const` is `emit_op(I32_CONST)` + `emit_leb_i32`,
    /// and `core_wasm::i32_const` does the same — so this cannot assert that
    /// the op is operand-less.
    pub fn emit_op(&mut self, op: Op, line: u32) {
        let bytes = op.encode();
        for b in bytes {
            self.emit(b, line);
        }
    }

    /// Emit `ref.null <heaptype>` — the spec instruction, immediate included.
    ///
    /// The heaptype is not decoration: `ref.null none` (a GC-heap null) traps
    /// on the GC accessors, while `ref.null extern` is the lenient null the
    /// dynamic languages use for JS `null` / PHP `NULL` / Python `None`. The VM
    /// used to express that difference with a second, custom opcode because
    /// `ref.null` had been declared with no immediate at all.
    pub fn emit_ref_null(&mut self, heaptype: u8, line: u32) {
        self.emit_op_u8(Op::NULL, heaptype, line);
    }

    /// Emit a `try_table` header with N catch clauses — the SINGLE SOURCE OF
    /// TRUTH for the VM's internal try_table byte layout. Every producer routes
    /// here: the shared `errors::emit_try_table`, the wast `WasmTryTable`
    /// lowering, and the `.wasm` reader importing foreign modules; the VM
    /// (`TRY_TABLE` dispatch) and the codec writer decode this exact layout.
    /// Layout: `[try_table, u8 clause_count, per clause: u8 kind, u16 tag(be),
    /// u16 offset(be)]`; clauses match by TAG IDENTITY in order. Each `(kind,
    /// tag)` uses the `CATCH_KIND_*` values (tag ignored for catch_all kinds).
    /// Returns each clause's offset-placeholder byte position — patch it with
    /// the forward distance to its handler once that handler is emitted.
    pub fn emit_try_table_clauses(&mut self, clauses: &[(u8, u16)], line: u32) -> Vec<usize> {
        self.emit_op(Op::TRY_TABLE, line);
        self.emit(clauses.len() as u8, line);
        let mut offset_positions = Vec::with_capacity(clauses.len());
        for &(kind, tag) in clauses {
            self.emit(kind, line);
            self.emit((tag >> 8) as u8, line);
            self.emit((tag & 0xff) as u8, line);
            offset_positions.push(self.current_offset());
            self.emit(0, line); // catch offset hi (placeholder)
            self.emit(0, line); // catch offset lo (placeholder)
        }
        offset_positions
    }

    pub fn emit_op_u16(&mut self, op: Op, operand: u16, line: u32) {
        self.emit_op(op, line);
        self.emit((operand >> 8) as u8, line);
        self.emit((operand & 0xff) as u8, line);
    }

    pub fn emit_op_u8(&mut self, op: Op, operand: u8, line: u32) {
        self.emit_op(op, line);
        self.emit(operand, line);
    }

    /// `array.new_fixed $t N` — BOTH immediates, in spec order.
    ///
    /// The type index is what stamps the array's rtt, and that stamp is what
    /// makes `array.get`/`set`/`fill`/`copy` bounds-check per the GC proposal.
    /// This instruction used to carry only `N`, so every array it built was
    /// indistinguishable from a dynamic one and could never trap. Pass `0` for
    /// a dynamic-language array literal, which is deliberately lenient.
    /// `struct.get` / `get_s` / `get_u` / `set` — `(typeidx, idx)`.
    /// typeidx 0 makes `idx` a constant-pool index for a field NAME; a real
    /// typeidx makes it a spec `fieldidx` into indexed storage.
    pub fn emit_struct_field_op(&mut self, op: Op, typeidx: u16, idx: u16, line: u32) {
        self.emit_op(op, line);
        self.emit((typeidx >> 8) as u8, line);
        self.emit((typeidx & 0xff) as u8, line);
        self.emit((idx >> 8) as u8, line);
        self.emit((idx & 0xff) as u8, line);
    }

    /// `struct.new $t N` — typeidx 0 means the dynamic object-literal form,
    /// where `count` is the number of key/value pairs on the stack. A real
    /// typeidx takes its field count from the type, per the GC spec.
    pub fn emit_struct_new(&mut self, typeidx: u16, count: u16, line: u32) {
        self.emit_op(Op::STRUCT_NEW, line);
        self.emit((typeidx >> 8) as u8, line);
        self.emit((typeidx & 0xff) as u8, line);
        self.emit((count >> 8) as u8, line);
        self.emit((count & 0xff) as u8, line);
    }

    pub fn emit_array_new_fixed(&mut self, typeidx: u16, count: u16, line: u32) {
        self.emit_op(Op::ARRAY_NEW_FIXED, line);
        self.emit((typeidx >> 8) as u8, line);
        self.emit((typeidx & 0xff) as u8, line);
        self.emit((count >> 8) as u8, line);
        self.emit((count & 0xff) as u8, line);
    }

    pub fn emit_op_u8_u8(&mut self, op: Op, first: u8, second: u8, line: u32) {
        self.emit_op(op, line);
        self.emit(first, line);
        self.emit(second, line);
    }

    pub fn emit_leb_u32(&mut self, mut value: u32, line: u32) {
        loop {
            let mut byte = (value & 0x7f) as u8;
            value >>= 7;
            if value != 0 {
                byte |= 0x80;
            }
            self.emit(byte, line);
            if value == 0 {
                break;
            }
        }
    }

    pub fn emit_leb_i32(&mut self, value: i32, line: u32) {
        let mut v = value as u32;
        loop {
            let mut byte = (v & 0x7f) as u8;
            v >>= 7;
            let sign_bit = (byte & 0x40) != 0;
            if (v == 0 && !sign_bit) || (v == 0xFFFF_FFFF >> 6 && sign_bit) {
                self.emit(byte, line);
                break;
            }
            byte |= 0x80;
            self.emit(byte, line);
        }
    }

    pub fn emit_leb_i64(&mut self, value: i64, line: u32) {
        let mut v = value as u64;
        loop {
            let mut byte = (v & 0x7f) as u8;
            v >>= 7;
            let sign_bit = (byte & 0x40) != 0;
            if (v == 0 && !sign_bit) || (v == u64::MAX >> 6 && sign_bit) {
                self.emit(byte, line);
                break;
            }
            byte |= 0x80;
            self.emit(byte, line);
        }
    }

    pub fn add_constant(&mut self, value: Value) -> u16 {
        self.constants.push(value);
        (self.constants.len() - 1) as u16
    }

    pub fn emit_jump(&mut self, op: Op, line: u32) -> usize {
        self.emit_op(op, line);
        self.emit(0xff, line);
        self.emit(0xff, line);
        self.code.len() - 2
    }

    pub fn patch_jump(&mut self, offset: usize) {
        let jump = self.code.len() as i32 - (offset as i32 + 2);
        self.code[offset] = (jump >> 8) as u8;
        self.code[offset + 1] = (jump & 0xff) as u8;
    }

    pub fn current_offset(&self) -> usize {
        self.code.len()
    }

    /// Get the source line number for a bytecode offset.
    /// Returns None if no line info is available for that offset.
    pub fn get_line(&self, offset: usize) -> Option<u32> {
        self.lines.get(offset).copied().filter(|&l| l > 0)
    }

    pub fn emit_loop(&mut self, target_offset: usize, line: u32) {
        self.emit_op(Op::BR, line);
        let jump = target_offset as i32 - (self.code.len() as i32 + 2);
        self.emit((jump >> 8) as u8, line);
        self.emit((jump & 0xff) as u8, line);
    }

    // ── Structured control flow (WASM-compliant) ───────────────────────
    //
    // BLOCK, LOOP, IF all carry a single blocktype byte (WASM §5.4.1):
    //   0x40 = void (no result), 0x6F = externref, 0x7F = i32, etc.
    //
    // The VM pre-scans each function body to build a block table mapping
    // every BLOCK/LOOP/IF/ELSE opcode position to its ELSE/END target ip.
    // No size headers or patching needed — nesting is the structure.

    /// Emit a void BLOCK (no value produced). The VM uses the pre-scanned
    /// block table to find the matching END; `patch_block` is a no-op.
    pub fn emit_block(&mut self, line: u32) -> usize {
        self.emit_block_typed(line, 0)
    }

    /// Emit BLOCK with explicit result count.
    /// 0 = void, 1 = single externref, 2+ = multi-value (WASM encoder registers type).
    pub fn emit_block_typed(&mut self, line: u32, result_count: u8) -> usize {
        self.emit_op(Op::BLOCK, line);
        self.emit(result_count, line); // raw count; WASM encoder translates to blocktype
        self.code.len() - 1 // dummy patch pos — patch_block is a no-op
    }

    /// Emit a void LOOP. Returns (dummy_patch, loop_body_start).
    /// `loop_body_start` is the ip right after the result_count byte —
    /// `br 0` inside the loop restarts there.
    pub fn emit_loop_s(&mut self, line: u32) -> (usize, usize) {
        self.emit_loop_typed(line, 0)
    }

    /// Emit LOOP with explicit result count.
    pub fn emit_loop_typed(&mut self, line: u32, result_count: u8) -> (usize, usize) {
        self.emit_op(Op::LOOP, line);
        self.emit(result_count, line);
        let dummy_patch = self.code.len() - 1;
        let loop_body_start = self.code.len();
        (dummy_patch, loop_body_start)
    }

    /// Close a BLOCK, LOOP, IF, or IF/ELSE.
    pub fn emit_end(&mut self, line: u32) {
        self.emit_op(Op::END, line);
    }

    /// No-op — block table replaces size-header patching.
    #[inline(always)]
    pub fn patch_block(&mut self, _patch_pos: usize) {}

    /// No-op — block table replaces size-header patching.
    #[inline(always)]
    pub fn patch_loop(&mut self, _patch_pos: usize) {}

    /// Emit IF that produces no value (void).
    /// Caller must have an i32 on the stack (non-zero = enter then-body).
    /// Use `emitter::ops::emit_dyn_to_bool` first to coerce a Value → i32.
    pub fn emit_if(&mut self, line: u32) -> usize {
        self.emit_op(Op::IF, line);
        self.emit(0u8, line); // result_count=0 → void
        self.code.len() - 1
    }

    /// Emit IF that leaves one externref on the stack (then/else both push one Value).
    pub fn emit_if_value(&mut self, line: u32) -> usize {
        self.emit_op(Op::IF, line);
        self.emit(1u8, line); // result_count=1 → externref
        self.code.len() - 1
    }

    /// Emit ELSE. Must be matched with emit_if + emit_end.
    pub fn emit_else(&mut self, line: u32) {
        self.emit_op(Op::ELSE, line);
    }

    /// Emit WASM `br` with a structured label depth.
    pub fn emit_br(&mut self, depth: u32, line: u32) {
        self.emit_op(Op::BR, line);
        self.emit_leb_u32(depth, line);
    }

    /// Emit WASM `br_if` with a structured label depth. Expects i32 on stack.
    pub fn emit_br_if(&mut self, depth: u32, line: u32) {
        self.emit_op(Op::BR_IF, line);
        self.emit_leb_u32(depth, line);
    }

    /// Emit WASM `br_table` with structured label depths.
    pub fn emit_br_table(&mut self, depths: &[u32], default_depth: u32, line: u32) {
        self.emit_op(Op::BR_TABLE, line);
        self.emit_leb_u32(depths.len() as u32, line);
        for &depth in depths {
            self.emit_leb_u32(depth, line);
        }
        self.emit_leb_u32(default_depth, line);
    }

    pub fn read_u16(&self, offset: usize) -> u16 {
        ((self.code[offset] as u16) << 8) | (self.code[offset + 1] as u16)
    }

    pub fn read_i16(&self, offset: usize) -> i16 {
        self.read_u16(offset) as i16
    }

    /// Allocate scratch locals for emitter helpers. Returns the base slot.
    /// Updates both `local_count` and `scratch_high_water` so the compiler
    /// always sees the correct total at finalization.
    pub fn alloc_scratch(&mut self, n: u16) -> u16 {
        let base = self.local_count;
        self.local_count += n;
        if self.local_count > self.scratch_high_water {
            self.scratch_high_water = self.local_count;
        }
        base
    }

    /// Finalize local_count after compilation. Ensures scratch allocations
    /// from emitter helpers are always included. Call this at the end of
    /// every function/lambda compilation path.
    pub fn finalize_local_count(&mut self, scope_next_slot: u16) {
        self.local_count = scope_next_slot
            .max(self.local_count)
            .max(self.scratch_high_water);
    }

    /// Emit CALL_IMPORT: [op, import_idx_hi, import_idx_lo, argc].
    pub fn emit_call(&mut self, import_idx: u16, argc: u8, line: u32) {
        self.emit_op(Op::CALL_IMPORT, line);
        self.emit((import_idx >> 8) as u8, line);
        self.emit((import_idx & 0xff) as u8, line);
        self.emit(argc, line);
    }

    /// Emit DUP via local.tee + local.get with a shared scratch local.
    pub fn emit_dup(&mut self, line: u32) {
        let slot = match self.dup_slot {
            Some(s) => s,
            None => {
                let s = self.alloc_scratch(1);
                self.dup_slot = Some(s);
                s
            }
        };
        self.emit_op_u16(Op::LOCAL_TEE, slot, line);
        self.emit_op_u16(Op::LOCAL_GET, slot, line);
    }

    /// Emit i32.const via signed LEB128.
    pub fn emit_i32_const(&mut self, value: i32, line: u32) {
        self.emit_op(Op::I32_CONST, line);
        self.emit_leb_i32(value, line);
    }

    /// Emit i64.const via signed LEB128.
    pub fn emit_i64_const(&mut self, value: i64, line: u32) {
        self.emit_op(Op::I64_CONST, line);
        self.emit_leb_i64(value, line);
    }

    /// Emit f32.const (4 raw LE bytes).
    pub fn emit_f32_const(&mut self, value: f32, line: u32) {
        self.emit_op(Op::F32_CONST, line);
        for b in value.to_le_bytes() {
            self.emit(b, line);
        }
    }

    /// Emit f64.const (8 raw LE bytes).
    pub fn emit_f64_const(&mut self, value: f64, line: u32) {
        self.emit_op(Op::F64_CONST, line);
        for b in value.to_le_bytes() {
            self.emit(b, line);
        }
    }

    /// Emit `ref.test` / `ref.cast` / their `_null` variants with a HEAPTYPE
    /// immediate, in the spec encoding (one signed LEB; negative = abstract,
    /// non-negative = module type index).
    pub fn emit_ref_type_op(&mut self, op: Op, ht: crate::opcode::heaptype::HeapType, line: u32) {
        self.emit_op(op, line);
        self.emit_leb_i32(ht.to_sleb(), line);
    }

    /// Emit a string constant — js-string-builtins § String constants:
    ///
    /// ```wasm
    /// (global (import "wasm:string-constants" "hello") (ref extern))
    /// global.get $that
    /// ```
    ///
    /// The import's field name IS the value, so nothing is called and nothing
    /// is encoded: the constant reaches the module through its own import.
    pub fn emit_string_const(&mut self, s: &str, line: u32) {
        self.add_global_import(STRING_CONSTANTS_MODULE, s);
        let key = imported_global_key(STRING_CONSTANTS_MODULE, s);
        let ci = self.intern_string_constant(&key);
        self.emit_op_u16(Op::GLOBAL_GET, ci, line);
    }

    /// Emit a boolean constant: i32.const N + call wasm:js-boolean.fromI32.
    pub fn emit_bool_const(&mut self, value: bool, line: u32) {
        self.emit_i32_const(if value { 1 } else { 0 }, line);
        let idx = self.add_import("wasm:js-boolean", "fromI32");
        self.emit_call(idx, 1, line);
    }
}

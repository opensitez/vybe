//! WASM type section encoding with Custom Descriptors.
//!
//! Each TypeEntry from the compiler produces TWO WASM GC types:
//! 1. The described struct type (object fields as externref)
//! 2. The descriptor struct type (JS prototype + vtable methods)
//!
//! This follows the Custom Descriptors proposal:
//! proposals/custom-descriptors/proposals/custom-descriptors/Overview.md

use crate::encoding::*;
use crate::writer::sections::HostImport;
use vybe_runtime::Chunk;

// Custom Descriptors binary encoding
const CD_DESCRIPTOR: u8 = 0x4D; // (descriptor $x) prefix
const CD_DESCRIBES: u8 = 0x4C; // (describes $x) prefix
const CD_SUB_FINAL: u8 = 0x4F; // sub final

/// Context for .wasm emission — maps internal types to WASM type indices.
pub struct WasmTypeContext {
    pub internal_callable_table: u32,
    pub table_operands: Vec<(u8, Option<u8>)>,
    /// Zero-argument source entry: (target func, adapter func, adapter type, results).
    pub source_entry: Option<(u32, u32, u32, u8)>,
    pub callables: super::callable::Plan,
    pub generators: Vec<super::generator_construct::Entry>,
    /// Exact diagnostic type name -> WASM index for a described struct.
    pub struct_type_indices: std::collections::HashMap<String, u32>,
    /// The module's own type index space → WASM type index. Slot `n - 1` is
    /// the 1-based immediate a `ref.test` / `struct.new` carries, so a type
    /// reference reaches the binary without passing through a name.
    pub struct_type_by_module_index: Vec<u32>,
    /// type_name → WASM type index for the descriptor struct (vtable + proto)
    pub desc_type_indices: std::collections::HashMap<String, u32>,
    /// Absolute WASM global index of the FIRST per-class descriptor singleton,
    /// or `None` when the module declares no classes.
    ///
    /// Custom Descriptors forbids `struct.new` / `struct.new_default` from
    /// allocating a type that carries a `(descriptor …)` clause — and
    /// `build_type_context` stamps one on every class — so each allocation
    /// needs a descriptor VALUE as its last operand. That value has to be one
    /// singleton per class, not one per allocation: `ref.cast_desc_eq` compares
    /// descriptors by IDENTITY, so per-allocation descriptors would make two
    /// instances of the same class fail to match each other, and each instance
    /// would reflect its own JS prototype out of descriptor field 0.
    ///
    /// The singletons are appended AFTER every module-defined global so that
    /// the index space `chunk::global_index_space` hands the compiler is
    /// untouched — class `i` (the same ordinal `desc_type_indices` is built
    /// from) owns `desc_global_base + i`.
    pub desc_global_base: Option<u32>,
    /// Per-class descriptor VTABLE plan, in class-ordinal order — one entry
    /// per descriptor global, holding the WASM function index of every method
    /// in `TypeEntry.methods` order.
    ///
    /// ⛔ Built inside the same loop that writes the descriptor struct's field
    /// count, because `struct.new` carries NO count immediate: the operand list
    /// **is** the field count. A plan assembled anywhere else can drift from
    /// `2 + te.methods.len()` by one and the module is not subtly wrong, it is
    /// invalid. The `1 +` / `2 +` split between writer and compiler is exactly
    /// that drift, caught only because nothing read a descriptor yet.
    ///
    /// `None` in a slot means the method's chunk index was out of range: that
    /// slot encodes `ref.null func` so the struct keeps its shape rather than
    /// the whole class falling back to a null vtable.
    pub desc_vtable: Vec<Vec<Option<u32>>>,
    /// type_name → vec of field names in order (for field index lookup)
    pub struct_fields: std::collections::HashMap<String, Vec<String>>,
    /// WASM type index for the dynamic array type — `(array (mut externref))`.
    /// Used for every `Value`-typed array in Vybe's uniform representation.
    pub array_type_idx: u32,
    /// WASM type index for `(array (mut i16))` — UTF-16 code-unit arrays.
    /// Strings-as-GC-array (inline string) path will reference this.
    pub string_array_type_idx: u32,
    /// WASM type index for `(array (mut i8))` — byte arrays. Backs
    /// `Uint8Array` / `Int8Array` when we inline TypedArrays.
    pub byte_array_type_idx: u32,
    /// First function type index (after GC types)
    pub func_type_base: u32,
    /// Total number of GC types (structs + descriptors + array types)
    pub gc_type_count: u32,
    /// arity → WASM type index for (externref * arity) -> externref
    pub func_type_by_arity: std::collections::HashMap<u8, u32>,
    /// Declared functype spelling → its type index. What `call_indirect` needs:
    /// arity cannot pick between two same-arity types.
    pub func_type_by_signature: std::collections::HashMap<String, u32>,
    /// Function indices whose DECLARED result is a bare `i32`/`f64` rather than
    /// the externref every other value carries, mapped to that valtype byte.
    ///
    /// ⛔ THE VALUE ABI IS EXTERNREF; THE SPEC SIGNATURES ARE NOT. A proposal
    /// builtin like `wasm:js-string.test` really does return i32, and once the
    /// import is typed truthfully its result can no longer be handed to an
    /// externref consumer unconverted — an `if (result externref)` arm ending
    /// in `call $length` is "expected externref, got i32". The call site has to
    /// box, so it has to know; this is how it knows.
    pub raw_result_funcs: std::collections::HashMap<u32, u8>,
    /// Declared result arity by absolute WASM function index.
    ///
    /// The dynamic VM value ABI historically treated host calls as producing a
    /// value, and bytecode commonly follows effect-only calls with `DROP`.
    /// Once proposal/host imports are typed truthfully, a `() -> ()` import
    /// followed by spec `drop` is invalid. The code writer uses this arity map
    /// to consume that legacy bytecode drop only for functions that really
    /// return nothing.
    pub func_result_arity: std::collections::HashMap<u32, u8>,

    /// Declared PARAM valtypes for imports that take a raw numeric operand.
    ///
    /// ⛔ THE MIRROR OF `raw_result_funcs`, AND IT WAS MISSING. A truthful
    /// signature has two halves and so does the ABI reconciliation: results
    /// are boxed onto the externref value ABI on the way out, and operands
    /// have to be unboxed on the way in. `wasm:js-string.fromF64` is declared
    /// `(param f64)` and every caller pushed a BOXED number, because this
    /// writer boxes every f64 it produces — so the argument arrived as an
    /// externref and the module was invalid at every single call site.
    /// Only imports with at least one non-externref param appear here.
    pub packed_i16_types: std::collections::HashMap<u16, u32>,
    pub import_ref_params: std::collections::HashMap<(u32, usize), (bool, i64)>,
    pub raw_param_funcs: std::collections::HashMap<u32, Vec<u8>>,
    /// Declared PARAM valtypes for imports. Dynamic VM calls may omit trailing
    /// optional args; wasm imports cannot, so the code writer pads that missing
    /// tail before emitting `call`.
    pub import_param_funcs: std::collections::HashMap<u32, Vec<u8>>,
    /// String-constant text → its GLOBAL index. Needed to put a property NAME
    /// on the stack for a dynamic (typeidx 0) property access, which lowers to
    /// a host call rather than a struct op. Computed here from the chunks so no
    /// call site has to thread it in.
    pub string_const_global: std::collections::HashMap<String, u32>,
    /// `externref^M -> externref^N` functype indices keyed by
    /// (param_count, result_count). Referenced by multi-value block
    /// headers as their s33 typeidx `blocktype`, AND by the
    /// call_indirect/return_call_indirect `(type $sig)` annotations
    /// (which must match the callee's functype exactly, result count
    /// included — first-seen-arity lookup is not exact).
    pub block_type_by_results: std::collections::HashMap<(u8, u8), u32>,
    /// `externref^M -> i32` functype indices for raw i32-result blocks that
    /// also take params. `(0 -> i32)` uses the one-byte blocktype shorthand;
    /// parametric forms need a real type index distinct from
    /// `block_type_by_results[(M, 1)]`, whose result is externref.
    pub block_i32_type_by_params: std::collections::HashMap<u8, u32>,
    /// Type index for `(externref) -> ()` — the shape required by the
    /// tag section's exception tag (exception-handling proposal).
    pub exception_type_idx: u32,
    /// Type index of the suspend/resume tag type
    /// `(tag (param externref) (result externref))` — used by the
    /// stack-switching proposal's `suspend` / `resume` when the
    /// module contains any `CONT_NEW` op. Zero if unused.
    pub suspend_tag_type_idx: u32,
    /// Continuation type index — `(cont $ft)` wrapping the shared
    /// single-arg single-result fiber function signature. Zero if
    /// the module doesn't use stack switching.
    pub continuation_type_idx: u32,
    pub continuation_carrier_type_idx: u32,
    pub continuation_landing_type_idx: u32,
    pub native_resume_type_idx: u32,
    pub native_carrier_resume_type_idx: u32,
    pub native_carrier_resume_function_idx: Option<u32>,
    pub native_resume_function_idx: Option<u32>,
    /// Continuation tag signature/name → emitted Wasm tag index.
    pub continuation_tag_indices: std::collections::HashMap<(String, String, String), u32>,
    /// Type indices for typed continuation tags, in tag-section order.
    pub continuation_tag_type_indices: Vec<u32>,
    /// Whether any `CONT_NEW` / `SUSPEND` / `RESUME` / `SWITCH` op
    /// was observed in the bytecode. Drives whether we emit the
    /// continuation type, the suspend tag, and the tag-section entry.
    pub uses_stack_switching: bool,
}

pub(crate) fn continuation_tag_key(
    tag: &vybe_runtime::chunk::ContinuationTag,
) -> (String, String, String) {
    (
        tag.name.clone(),
        tag.yield_type.clone(),
        tag.resume_type.clone(),
    )
}

pub(crate) fn continuation_tag_valtype(type_name: &str) -> u8 {
    match type_name.to_ascii_lowercase().as_str() {
        "i32" | "bool" | "boolean" => TYPE_I32,
        "i64" => TYPE_I64,
        "f32" => TYPE_F32,
        "f64" | "number" => TYPE_F64,
        _ => TYPE_EXTERNREF,
    }
}

fn collect_continuation_tags(chunks: &[Chunk]) -> Vec<vybe_runtime::chunk::ContinuationTag> {
    let mut seen = std::collections::HashSet::new();
    let mut tags = Vec::new();
    for chunk in chunks {
        for tag in &chunk.continuation_tags {
            let key = continuation_tag_key(tag);
            if seen.insert(key) {
                tags.push(tag.clone());
            }
        }
    }
    tags
}

impl WasmTypeContext {
    /// Look up the WASM type index for a described struct type by name.
    pub fn struct_type(&self, name: &str) -> Option<u32> {
        self.struct_type_indices.get(name).copied()
    }

    /// Look up the WASM type index for a descriptor type by name.
    /// The WASM type index for a 1-based module type index — the immediate
    /// form, no name involved.
    pub fn struct_type_by_index(&self, module_index: u32) -> Option<u32> {
        if module_index == 0 {
            return None;
        }
        self.struct_type_by_module_index
            .get(module_index as usize - 1)
            .copied()
    }

    pub fn desc_type(&self, name: &str) -> Option<u32> {
        self.desc_type_indices.get(name).copied()
    }

    /// How many per-class descriptor singletons this module declares — one per
    /// class, so one per described/descriptor pair.
    pub fn descriptor_global_count(&self) -> u32 {
        // Names are diagnostic lookup aliases, not type identities. Distinct
        // module type rows may carry the same name; their descriptor globals
        // must still match the structural class-ordinal/vtable plan.
        self.desc_vtable.len() as u32
    }

    /// The global holding class `name`'s descriptor singleton.
    ///
    /// Derived from the descriptor TYPE index rather than a second map, so the
    /// two cannot drift: `build_type_context` lays each class out as the pair
    /// `(2i, 2i+1)`, so `desc_type_indices[name] == 2i + 1` and the class
    /// ordinal is `(idx - 1) / 2`.
    pub fn desc_global(&self, name: &str) -> Option<u32> {
        let base = self.desc_global_base?;
        let desc_idx = self.desc_type(name)?;
        Some(base + (desc_idx - 1) / 2)
    }

    /// The descriptor singleton for a class given its DESCRIBED wasm type
    /// index — the index that actually reaches the binary.
    ///
    /// Use this when the type was resolved rather than named: the dynamic
    /// `struct.new` path guesses a wasm type from the field count and never
    /// holds a module index, but the guess can still land on a
    /// descriptor-carrying class. Described types are the even members of the
    /// `(2i, 2i+1)` layout, so an odd index (a descriptor struct, which has no
    /// descriptor of its own) or one past the pairs correctly answers `None`.
    pub fn desc_global_for_described(&self, described_idx: u32) -> Option<u32> {
        let base = self.desc_global_base?;
        if described_idx % 2 != 0 {
            return None;
        }
        let ordinal = described_idx / 2;
        if ordinal >= self.array_type_idx / 2 {
            return None;
        }
        Some(base + ordinal)
    }

    /// The descriptor singleton for a 1-based MODULE type index.
    ///
    /// ⚠ Prefer this over the by-name form inside the code encoder. The type
    /// table lives on `chunks[0]` ONLY, and `encode_code_section` walks every
    /// chunk, so a name lookup through `chunk.types` silently misses in every
    /// chunk but the first — which is where constructors live.
    pub fn desc_global_by_index(&self, module_index: u32) -> Option<u32> {
        let described = self.struct_type_by_index(module_index)?;
        self.desc_global_for_described(described)
    }

    /// Look up the field index for a field name within a struct type.
    pub fn field_index(&self, type_name: &str, field_name: &str) -> Option<u32> {
        let fields = self.struct_fields.get(type_name)?;
        fields
            .iter()
            .position(|f| f == field_name)
            .map(|i| i as u32)
    }
}

/// A declared value type's binary encoding. Anything unrecognised — a GC
/// reference spelling this table does not carry — falls back to `externref`,
/// which is what the whole section used to be, so an unknown type degrades to
/// today's behaviour rather than emitting a wrong byte.
pub(super) fn val_type_byte(spelling: &str) -> u8 {
    match spelling.trim() {
        "i32" => 0x7F,
        "i64" => 0x7E,
        "f32" => 0x7D,
        "f64" => 0x7C,
        "v128" => 0x7B,
        "nullexnref" => 0x74,
        "nullfuncref" => 0x73,
        "nullexternref" => 0x72,
        "nullref" => 0x71,
        "funcref" => 0x70,
        "anyref" => 0x6E,
        "eqref" => 0x6D,
        "i31ref" => 0x6C,
        "structref" => 0x6B,
        "arrayref" => 0x6A,
        "exnref" => 0x69,
        _ => TYPE_EXTERNREF,
    }
}

/// Build the type context and encode the type section.
/// Layout: [rec group: (described struct + descriptor struct) per TypeEntry] [array type] [function types]
/// Ask the owning proposal module for a `(module, name)` pair's exact WASM
/// signature. Returns `false` when no proposal claims the pair, leaving the
/// caller to supply its own fallback — the `TYPE_FUNC` tag byte has already
/// been pushed either way.
///
/// ⛔ BOTH import tables must ask this. A builtin is the same function whether
/// it arrives as a host import or a runtime import; typing one table by arity
/// and the other by signature makes the SAME callee two different types.
fn write_proposal_signature(out: &mut Vec<u8>, module: &str, name: &str, chars_type: u32) -> bool {
    // These exact WIT scalar results lower to core i64. Leave their raw bits
    // on the stack for standard unsigned conversion, not an externref ABI.
    if matches!((module, name),
        ("wasi:random/insecure", "get-insecure-random-u64")
        | ("wasi:random/random", "get-random-u64")) {
        write_leb128_u32(out, 0);
        write_leb128_u32(out, 1);
        out.push(TYPE_I64);
        true
    } else if module == "ecma:function" && name == "apply" {
        write_leb128_u32(out, 3);
        out.extend([TYPE_EXTERNREF; 3]);
        write_leb128_u32(out, 1);
        out.push(TYPE_EXTERNREF);
        true
    } else if module == "wasm:text-encoder" && name == "measureStringAsUTF8" {
        // js-string-builtins UTF-8 extension: (externref) -> i32.
        write_leb128_u32(out, 1);
        out.push(TYPE_EXTERNREF);
        write_leb128_u32(out, 1);
        out.push(TYPE_I32);
        true
    } else if module == crate::writer::builtins::canon_builtins::MODULE {
        crate::writer::builtins::canon_builtins::write_signature(out, name)
    } else if module == crate::writer::builtins::js_string_builtins::MODULE {
        crate::writer::builtins::js_string_builtins::write_signature(out, name, chars_type)
    } else if module == crate::writer::builtins::js_array_builtins::MODULE {
        crate::writer::builtins::js_array_builtins::write_signature(out, name)
    } else if module == crate::writer::builtins::js_object_builtins::MODULE {
        crate::writer::builtins::js_object_builtins::write_signature(out, name)
    } else if module == crate::writer::builtins::js_map_builtins::MODULE {
        crate::writer::builtins::js_map_builtins::write_signature(out, name)
    } else if module == crate::writer::builtins::js_set_builtins::MODULE {
        crate::writer::builtins::js_set_builtins::write_signature(out, name)
    } else if module == crate::writer::builtins::js_json_builtins::MODULE {
        crate::writer::builtins::js_json_builtins::write_signature(out, name)
    } else if module == crate::writer::builtins::js_fixedarray_builtins::MODULE {
        crate::writer::builtins::js_fixedarray_builtins::write_signature(out, name)
    } else if module == crate::writer::builtins::js_structured_clone::MODULE {
        crate::writer::builtins::js_structured_clone::write_signature(out, name)
    } else if module == crate::writer::builtins::js_arraybuffer_builtins::ARRAYBUFFER_MODULE {
        crate::writer::builtins::js_arraybuffer_builtins::write_arraybuffer_signature(out, name)
    } else if module == crate::writer::builtins::js_arraybuffer_builtins::SHAREDARRAYBUFFER_MODULE {
        crate::writer::builtins::js_arraybuffer_builtins::write_sharedarraybuffer_signature(
            out, name,
        )
    } else if module == crate::writer::builtins::js_arraybuffer_builtins::DATAVIEW_MODULE {
        crate::writer::builtins::js_arraybuffer_builtins::write_dataview_signature(out, name)
    } else if module == crate::writer::builtins::js_weakmap_builtins::WEAKMAP_MODULE {
        crate::writer::builtins::js_weakmap_builtins::write_weakmap_signature(out, name)
    } else if module == crate::writer::builtins::js_weakmap_builtins::WEAKSET_MODULE {
        crate::writer::builtins::js_weakmap_builtins::write_weakset_signature(out, name)
    } else if let Some(variant) = crate::writer::builtins::js_typedarray_builtins::VARIANTS
        .iter()
        .copied()
        .find(|variant| variant.module() == module)
    {
        crate::writer::builtins::js_typedarray_builtins::write_signature(out, variant, name)
    } else if crate::writer::builtins::web_builtins::write_signature(out, module, name) {
        true
    } else {
        crate::writer::builtins::js_primitive_builtins::write_signature(out, module, name)
    }
}

pub(crate) fn has_proposal_signature(module: &str, name: &str) -> bool {
    if crate::writer::builtins::web_builtins::has_overloaded_signature(module, name) {
        return false;
    }
    write_proposal_signature(&mut Vec::new(), module, name, 0)
}

/// Decode the param valtypes of the functype just appended at `start`
/// (the byte after the `TYPE_FUNC` tag). Single-byte counts are not assumed.
fn read_native_valtype(out: &[u8], i: &mut usize) -> Option<(u8, Option<(bool, i64)>)> {
    let tag = *out.get(*i)?;
    *i += 1;
    if tag != 0x63 && tag != 0x64 {
        return Some((tag, None));
    }
    let (mut n, mut shift) = (0i64, 0u32);
    loop {
        let byte = *out.get(*i)?;
        *i += 1;
        n |= i64::from(byte & 127) << shift;
        shift += 7;
        if byte & 128 == 0 {
            if byte & 64 != 0 {
                n |= !0i64 << shift;
            }
            return Some((tag, Some((tag == 0x63, n))));
        }
        if shift >= 35 {
            return None;
        }
    }
}
fn signature_parts(out: &[u8], start: usize) -> Option<(Vec<(u8, Option<(bool, i64)>)>, Vec<u8>)> {
    let mut i = start;
    let n = read_leb_u32_from(out, &mut i)?;
    let mut params = Vec::new();
    for _ in 0..n {
        params.push(read_native_valtype(out, &mut i)?);
    }
    let n = read_leb_u32_from(out, &mut i)?;
    let mut results = Vec::new();
    for _ in 0..n {
        results.push(read_native_valtype(out, &mut i)?.0);
    }
    Some((params, results))
}
fn signature_params(out: &[u8], start: usize) -> Vec<u8> {
    signature_parts(out, start).map_or_else(Vec::new, |(p, _)| p.into_iter().map(|v| v.0).collect())
}

fn read_leb_u32_from(out: &[u8], i: &mut usize) -> Option<u32> {
    let mut value = 0u32;
    let mut shift = 0u32;
    loop {
        let b = *out.get(*i)?;
        *i += 1;
        value |= ((b & 0x7f) as u32) << shift;
        if b & 0x80 == 0 {
            return Some(value);
        }
        shift += 7;
        if shift >= 35 {
            return None;
        }
    }
}

/// Decode the result valtypes of the functype just appended at `start`
/// (the byte after the `TYPE_FUNC` tag).
fn signature_results(out: &[u8], start: usize) -> Vec<u8> {
    signature_parts(out, start).map_or_else(Vec::new, |(_, r)| r)
}

fn record_result_signature(ctx: &mut WasmTypeContext, func_idx: u32, out: &[u8], sig_start: usize) {
    let results = signature_results(out, sig_start);
    ctx.func_result_arity.insert(func_idx, results.len() as u8);
    if results.len() == 1 {
        let vt = results[0];
        if vt == TYPE_I32 || vt == TYPE_F32 || vt == TYPE_F64 {
            ctx.raw_result_funcs.insert(func_idx, vt);
        }
    }
}

/// Record an import whose declared params are not all `externref`, so the
/// `call` site can unbox them. i64 params are DELIBERATELY not recorded:
/// there is no i64 unbox — our value ABI is externref and js-primitive-builtins
/// exposes only `wasm:js-bigint.test`, the i64 conversions having been removed
/// from the proposal. Recording one would produce a call site that unboxes to
/// i32 and silently truncates. `wasm:js-string.fromI64`/`fromU64` are the only
/// two, and no emitter in the tree calls either.
fn record_raw_params(ctx: &mut WasmTypeContext, func_idx: u32, out: &[u8], sig_start: usize) {
    if let Some((params, _)) = signature_parts(out, sig_start) {
        for (index, (_, r)) in params.into_iter().enumerate() {
            if let Some(r) = r {
                ctx.import_ref_params.insert((func_idx, index), r);
            }
        }
    }
    let params = signature_params(out, sig_start);
    ctx.import_param_funcs.insert(func_idx, params.clone());
    if params.iter().any(|&t| t == TYPE_I64) {
        return;
    }
    if params.iter().any(|&t| t != TYPE_EXTERNREF) {
        ctx.raw_param_funcs.insert(func_idx, params);
    }
}

/// Order binary class pairs without changing the module's declared type IDs.
/// Recursive field/descriptor links may point forward; superclass links may not.
pub(crate) fn is_standard_mutable_i16(row: &vybe_runtime::chunk::TypeEntry) -> bool {
    row.kind == vybe_runtime::chunk::CompositeKind::Array
        && row.fields == ["i16"]
        && row.parent_index == 0
        && row.describes_index == 0
        && row.methods.is_empty()
        && row.implements.is_empty()
        && row.constructor_chunk.is_none()
        && !row.is_interface
        && row.field_descriptors.get("i16").is_none_or(|d| d.writable)
}
fn ordered_class_rows(rows: &[vybe_runtime::chunk::TypeEntry]) -> Result<Vec<usize>, String> {
    let mut children = vec![Vec::new(); rows.len()];
    let mut indegree = vec![0usize; rows.len()];
    let mut ready = std::collections::BTreeSet::new();
    let mut class_count = 0usize;
    for (i, row) in rows.iter().enumerate() {
        if is_standard_mutable_i16(row) { continue; }
        if row.describes_index != 0 {
            let described = row.describes_index as usize - 1;
            if rows.get(described).is_none_or(|r| r.describes_index != 0 || is_standard_mutable_i16(r)) {
                return Err(format!(
                    "descriptor type #{} '{}' describes invalid class #{}",
                    i + 1,
                    row.name,
                    row.describes_index
                ));
            }
            continue;
        }
        class_count += 1;
        if row.parent_index == 0 {
            ready.insert(i);
        } else {
            let parent = row.parent_index as usize - 1;
            let Some(parent_row) = rows.get(parent) else {
                return Err(format!(
                    "class type #{} '{}' has out-of-range parent #{}",
                    i + 1,
                    row.name,
                    row.parent_index
                ));
            };
            if parent_row.describes_index != 0 || is_standard_mutable_i16(parent_row) {
                return Err(format!(
                    "class type #{} '{}' has descriptor parent #{}",
                    i + 1,
                    row.name,
                    row.parent_index
                ));
            }
            children[parent].push(i);
            indegree[i] = 1;
        }
    }
    let mut ordered = Vec::with_capacity(class_count);
    while let Some(parent) = ready.pop_first() {
        ordered.push(parent);
        for &child in &children[parent] {
            indegree[child] -= 1;
            if indegree[child] == 0 {
                ready.insert(child);
            }
        }
    }
    if ordered.len() != class_count {
        let pending = rows
            .iter()
            .enumerate()
            .filter(|(i, r)| r.describes_index == 0 && indegree[*i] != 0)
            .map(|(i, r)| format!("#{} '{}'", i + 1, r.name))
            .collect::<Vec<_>>()
            .join(", ");
        return Err(format!("cycle in class supertypes: {pending}"));
    }
    Ok(ordered)
}

pub fn build_type_context(
    chunks: &[Chunk],
    host_imports: &[HostImport],
    import_count: usize,
    rt_imports: &[(&str, &str)],
) -> (Vec<u8>, WasmTypeContext) {
    let mut out = Vec::new();
    // Global index order is `rt_globals()` first, then the string constants —
    // see `sections::encode_import_section`. Computed from the chunks so the
    // map costs no call site a new parameter.
    let string_const_base = crate::writer::sections::rt_globals().len() as u32;
    let string_const_global: std::collections::HashMap<String, u32> =
        crate::writer::sections::collect_string_constants(chunks)
            .into_iter()
            .enumerate()
            .map(|(i, text)| (text, string_const_base + i as u32))
            .collect();
    let mut ctx = WasmTypeContext {
        internal_callable_table: chunks.first().map_or(0, |chunk| chunk.table_min_sizes.len() as u32),
        table_operands: chunks.first().map_or_else(Vec::new, |chunk| {
            (0..chunk.table_min_sizes.len()).map(|index| {
                let address = if chunk.table_is_64.get(index).copied().unwrap_or(false) {
                    TYPE_I64
                } else { TYPE_I32 };
                let reference = chunk.table_element_types.get(index).and_then(|entry| {
                    match entry.0 {
                        vybe_runtime::opcode::heaptype::HeapType::Abstract(0x70) => Some(0x70),
                        vybe_runtime::opcode::heaptype::HeapType::Abstract(0x6f) => Some(0x6f),
                        _ => None,
                    }
                });
                (address, reference)
            }).collect()
        }),
        source_entry: None,
        callables: super::callable::Plan::default(),
        generators: Vec::new(),
        struct_type_indices: std::collections::HashMap::new(),
        struct_type_by_module_index: Vec::new(),
        desc_type_indices: std::collections::HashMap::new(),
        desc_global_base: None,
        desc_vtable: Vec::new(),
        struct_fields: std::collections::HashMap::new(),
        array_type_idx: 0,
        string_array_type_idx: 0,
        byte_array_type_idx: 0,
        func_type_base: 0,
        gc_type_count: 0,
        func_type_by_arity: std::collections::HashMap::new(),
        func_type_by_signature: std::collections::HashMap::new(),
        raw_result_funcs: std::collections::HashMap::new(),
        func_result_arity: std::collections::HashMap::new(),
        packed_i16_types: std::collections::HashMap::new(),
        import_ref_params: std::collections::HashMap::new(),
        raw_param_funcs: std::collections::HashMap::new(),
        import_param_funcs: std::collections::HashMap::new(),
        string_const_global,
        block_type_by_results: std::collections::HashMap::new(),
        block_i32_type_by_params: std::collections::HashMap::new(),
        exception_type_idx: 0,
        suspend_tag_type_idx: 0,
        continuation_type_idx: 0,
        continuation_carrier_type_idx: 0,
        continuation_landing_type_idx: 0,
        native_resume_type_idx: 0,
        native_carrier_resume_type_idx: 0,
        native_carrier_resume_function_idx: None,
        native_resume_function_idx: None,
        continuation_tag_indices: std::collections::HashMap::new(),
        continuation_tag_type_indices: Vec::new(),
        uses_stack_switching: false,
    };

    // Collect TypeEntry definitions from chunk 0.
    //
    // ⛔ DESCRIPTOR ROWS ARE NOT CLASSES — skip them here.
    //
    // Each entry below emits a `(described 2i, descriptor 2i+1)` PAIR, so a
    // row that IS a descriptor would get a pair of its own: `#desc base`
    // acquiring `#desc base__desc`, doubling the type section and leaving the
    // real descriptor unreferenced. `describes_index` is a back-pointer (0 for
    // an ordinary class row), which is what makes the test exact rather than a
    // guess at the row's name.
    let all_rows = chunks.first().map(|c| c.types.as_slice()).unwrap_or(&[]);
    let class_rows = ordered_class_rows(all_rows)
        .unwrap_or_else(|error| panic!("invalid WASM type hierarchy: {error}"));
    let type_entries: Vec<_> = class_rows.iter().map(|&row| &all_rows[row]).collect();

    // Layout:
    // One big rec group holding every described/descriptor pair so that
    // described↔descriptor back-pointers resolve recursively. Class pairs
    // are ordered parent-first because supertype links cannot point forward.
    // Then: 3 array types (each its own implicit singleton rec group)
    // Then: function types for imports + chunks
    // Then: N multi-value block types (one per distinct block `result_count >= 2`)
    // Then: 1 exception type `(externref) -> ()` for the tag section
    let gc_struct_pairs = type_entries.len() as u32;
    let array_count = 3u32;
    let func_count = (import_count + chunks.len()) as u32;
    let entries = super::generator_construct::entries(chunks);
    let old_gc_count = gc_struct_pairs * 2 + array_count;
    let gc_type_count = old_gc_count + 3;
    ctx.callables.cell = old_gc_count;
    ctx.callables.captures = old_gc_count + 1;
    ctx.callables.carrier = old_gc_count + 2;
    ctx.callables.import_count = import_count as u32;
    ctx.callables.locals = chunks
        .iter()
        .map(|c| {
            (
                c as *const Chunk as usize,
                (
                    super::code::total_temp_locals(c),
                    super::callable::captured_slots(c),
                ),
            )
        })
        .collect();
    ctx.callables.chunks = chunks
        .iter()
        .map(|c| (c.arity, super::callable::source(c)))
        .collect();
    ctx.gc_type_count = gc_type_count;
    ctx.array_type_idx = gc_struct_pairs * 2;
    ctx.string_array_type_idx = gc_struct_pairs * 2 + 1;
    for (i, row) in all_rows.iter().enumerate() {
        if is_standard_mutable_i16(row) {
            ctx.packed_i16_types.insert(u16::try_from(i + 1).expect("declared packed array index fits u16"), ctx.string_array_type_idx);
            ctx.struct_type_indices.insert(row.name.clone(), ctx.string_array_type_idx);
        }
    }
    ctx.byte_array_type_idx = gc_struct_pairs * 2 + 2;
    ctx.func_type_base = gc_type_count;

    // Pre-scan chunks for BLOCK/LOOP/IF blocktypes needing a typeidx: any
    // (param_count, result_count) pair beyond the single-byte spec forms
    // (0,0)=void and (0,1)=one valtype. Each distinct pair gets its own
    // `externref^M -> externref^N` function type, referenced as an s33
    // typeidx blocktype at emission.
    let mut block_result_counts: std::collections::BTreeSet<(u8, u8)> =
        std::collections::BTreeSet::new();
    let mut source_call_counts: std::collections::BTreeSet<(u8, u8)> =
        std::collections::BTreeSet::new();
    let mut block_i32_param_counts: std::collections::BTreeSet<u8> =
        std::collections::BTreeSet::new();
    // A tag's type is `externref^arity -> ()`, the same shape a blocktype of
    // (arity, 0) has — so declaring one here gives the tag section a typeidx to
    // point at. Without this every tag had to borrow the single one-param
    // exception type, which is why a 2-ary tag could not be expressed at all.
    for chunk in chunks {
        for tag in &chunk.tags {
            block_result_counts.insert((tag.arity, 0));
        }
    }
    for chunk in chunks {
        let code = &chunk.code;
        let mut bip = 0;
        while bip + 3 < code.len() {
            let g = ((code[bip] as u16) << 8) | code[bip + 1] as u16;
            let s = ((code[bip + 2] as u16) << 8) | code[bip + 3] as u16;
            if let Some(op) = vybe_runtime::opcode::Op::decode(g, s) {
                if op == vybe_runtime::opcode::Op::BLOCK
                    || op == vybe_runtime::opcode::Op::LOOP
                    || op == vybe_runtime::opcode::Op::IF
                {
                    // Layout: group (2) + sub (2) + params (1) + results (1).
                    if bip + 5 < code.len() {
                        let (params, results) =
                            crate::writer::code::effective_structured_block_counts(chunk, op, bip);
                        if params > 0 || results >= 2 {
                            block_result_counts.insert((params, results));
                        }
                        if results == 1 && params > 0 && chunk.block_i32_results.contains(&bip) {
                            block_i32_param_counts.insert(params);
                        }
                    }
                } else if op == vybe_runtime::opcode::Op::CALL_INDIRECT
                    || op == vybe_runtime::opcode::Op::RETURN_CALL_INDIRECT
                {
                    // Immediates: argc (1) + tableidx (1) + results (1).
                    // The call's `(type $sig)` annotation must match the
                    // callee's functype EXACTLY — including the result
                    // count — so register the (params, results) pair.
                    // Unlike blocktypes, (0,0)/(0,1) have no shorthand here.
                    if bip + 6 < code.len() {
                        block_result_counts.insert((code[bip + 4], code[bip + 6]));
                    }
                } else if op == vybe_runtime::opcode::Op::CALL_REF
                    || op == vybe_runtime::opcode::Op::RETURN_CALL
                    || op == vybe_runtime::opcode::Op::RETURN_CALL_REF
                {
                    // Immediates: argc (1) + results (1). Compilers emit
                    // results=1 (uniform boxed ABI); reader-ingested calls
                    // carry their functype's true result count.
                    if bip + 5 < code.len() {
                        block_result_counts.insert((code[bip + 4], code[bip + 5]));
                        if super::callable::source(chunk)
                            || (op != vybe_runtime::opcode::Op::RETURN_CALL
                                && super::callable::uses_carrier_call(chunk, bip))
                        {
                            source_call_counts.insert((code[bip + 4], code[bip + 5]));
                        }
                    }
                }
                bip += crate::writer::code::opcode_size(op, code, bip);
            } else {
                bip += 4;
            }
        }
    }
    for c in chunks {
        if super::callable::source(c) {
            source_call_counts.insert((c.arity, c.result_arity.max(1)));
        }
    }
    for e in &entries {
        source_call_counts.insert((e.user_arity, 1));
    }
    let block_type_count = block_result_counts.len() as u32
        + block_i32_param_counts.len() as u32
        + source_call_counts.len() as u32;

    // Pre-scan for stack-switching usage. Any CONT_NEW/SUSPEND/RESUME/
    // SWITCH opcode triggers the emission of:
    //   * one suspend tag type `(func (param externref) (result externref))`
    //   * one continuation type `(cont <suspend-tag-func-type>)`
    //   * the matching tag section entry
    let uses_stack_switching = chunks.iter().any(|chunk| {
        let code = &chunk.code;
        let mut bip = 0;
        while bip + 3 < code.len() {
            let g = ((code[bip] as u16) << 8) | code[bip + 1] as u16;
            let s = ((code[bip + 2] as u16) << 8) | code[bip + 3] as u16;
            if let Some(op) = vybe_runtime::opcode::Op::decode(g, s) {
                if matches!(op,
                    o if o == vybe_runtime::opcode::Op::CONT_NEW
                      || o == vybe_runtime::opcode::Op::CONT_BIND
                      || o == vybe_runtime::opcode::Op::SUSPEND
                      || o == vybe_runtime::opcode::Op::RESUME
                      || o == vybe_runtime::opcode::Op::RESUME_THROW
                      || o == vybe_runtime::opcode::Op::SWITCH
                ) {
                    return true;
                }
                bip += crate::writer::code::opcode_size(op, code, bip);
            } else {
                bip += 4;
            }
        }
        false
    });
    ctx.uses_stack_switching = uses_stack_switching;
    let continuation_tags = collect_continuation_tags(chunks);
    let typed_continuation_type_count = continuation_tags.len() as u32;
    let native_next = chunks.iter().any(crate::writer::code::has_native_next);
    let stack_switching_type_count: u32 = if uses_stack_switching {
        4 + 2 * u32::from(native_next) + typed_continuation_type_count
    } else {
        typed_continuation_type_count
    };

    // Index layout: [gc] [func] [block] [suspend_tag_func] [continuation] [typed continuation tag funcs] [exception]
    let ss_base = gc_type_count + func_count + block_type_count;
    if uses_stack_switching {
        ctx.suspend_tag_type_idx = ss_base;
        ctx.continuation_type_idx = ss_base + 1;
        ctx.continuation_carrier_type_idx = ss_base + 2;
        ctx.continuation_landing_type_idx = ss_base + 3;
        ctx.native_resume_type_idx = ss_base + 4;
        ctx.native_carrier_resume_type_idx = ss_base + 5;
        ctx.native_carrier_resume_function_idx = native_next.then_some((import_count + chunks.len() + 1) as u32);
        ctx.native_resume_function_idx =
            native_next.then_some((import_count + chunks.len()) as u32);
    }
    let typed_tag_type_base = ss_base
        + if uses_stack_switching {
            4 + 2 * u32::from(native_next)
        } else {
            0
        };
    for (i, tag) in continuation_tags.iter().enumerate() {
        ctx.continuation_tag_indices
            .insert(continuation_tag_key(tag), 2 + i as u32);
        ctx.continuation_tag_type_indices
            .push(typed_tag_type_base + i as u32);
    }
    ctx.exception_type_idx = ss_base + stack_switching_type_count;
    let generator_type_base = ctx.exception_type_idx + 1;
    ctx.generators = entries
        .into_iter()
        .enumerate()
        .map(|(i, mut e)| {
            e.cont_type = generator_type_base + i as u32;
            e.constructor_func =
                (import_count + chunks.len()) as u32 + 2 * u32::from(native_next) + i as u32;
            e.body_func = (import_count + e.chunk_index) as u32;
            e.body_type = ctx.func_type_base + (import_count + e.chunk_index) as u32;
            e
        })
        .collect();
    ctx.callables.targets = (0..chunks.len())
        .map(|i| {
            ctx.generators
                .iter()
                .find(|e| e.chunk_index == i)
                .map_or((import_count + i) as u32, |e| e.constructor_func)
        })
        .collect();
    for e in &ctx.generators {
        ctx.callables.chunks[e.chunk_index].0 = e.user_arity;
    }
    let start_chunk = super::sections::start_chunk_index(chunks);
    if let Some(entry) = chunks.get(start_chunk)
        && super::callable::source(entry)
        && entry.arity == 0
        && !ctx.generators.iter().any(|entry| entry.chunk_index == start_chunk)
    {
        ctx.source_entry = Some((
            (import_count + start_chunk) as u32,
            (import_count + chunks.len()) as u32
                + 2 * u32::from(native_next) + ctx.generators.len() as u32,
            generator_type_base + ctx.generators.len() as u32,
            entry.result_arity.max(1),
        ));
    }

    // Populate maps from the binary parent-first order while retaining every
    // original module type ID, including interleaved explicit descriptor rows.
    // ⚠ `struct_type_by_module_index` is keyed by MODULE index, and the module
    // index space includes the descriptor rows that `type_entries` filtered
    // out. Building it from the filtered list would shift every entry past the
    // first descriptor row, so it is built separately over ALL rows below —
    // a class row answers with its described index, a descriptor row with its
    // paired `2i+1`, resolved through the `describes_index` back-pointer.
    {
        let mut class_ordinal: Vec<Option<u32>> = vec![None; all_rows.len()];
        for (ordinal, &module_row) in class_rows.iter().enumerate() {
            class_ordinal[module_row] = Some(ordinal as u32);
        }
        ctx.struct_type_by_module_index = all_rows
            .iter()
            .enumerate()
            .map(|(i, t)| {
                if is_standard_mutable_i16(t) {
                    ctx.string_array_type_idx
                } else if t.describes_index == 0 {
                    class_ordinal[i].map_or(0, |k| k * 2)
                } else {
                    // `describes_index` is 1-BASED into the same table.
                    class_ordinal
                        .get(t.describes_index as usize - 1)
                        .copied()
                        .flatten()
                        .map_or(0, |k| k * 2 + 1)
                }
            })
            .collect();
        // A descriptor row is nameable too — `struct.new #desc <C>` has to
        // resolve to the descriptor struct, not to nothing.
        for (i, t) in all_rows.iter().enumerate() {
            if t.describes_index != 0 {
                if let Some(&idx) = ctx.struct_type_by_module_index.get(i) {
                    ctx.struct_type_indices.insert(t.name.clone(), idx);
                    ctx.struct_fields.insert(t.name.clone(), t.fields.clone());
                }
            }
        }
    }
    for (i, te) in type_entries.iter().enumerate() {
        let described_idx = (i as u32) * 2;
        let descriptor_idx = (i as u32) * 2 + 1;
        // ⛔ NOT `to_lowercase()`. These maps used to fold the key
        // unconditionally, which made `class Foo` and `class FOO` ONE entry:
        // the type section declared two described/descriptor pairs while the
        // global section emitted ONE singleton, `FOO`'s descriptor type had no
        // singleton at all, and `desc_global("FOO")` handed back `Foo`'s global
        // typed `(ref (exact 1))` — an INVALID module, and under the proposal
        // `ref.cast_desc_eq` would also test the two classes as each other.
        //
        // It is the "a QUERY folded UPSTREAM defeats it" shape from
        // `casesensitivityplan.md`, at a WRITE site, where it cannot be undone.
        // The conditional fold could never reach it because the fold happened at
        // the KEY. Nothing is lost by removing it: the compiler's `canon()`
        // already folds names for a case-insensitive language BEFORE they get
        // here, so those keys arrive uniform, while a case-sensitive language
        // keeps the distinction its source made. Found by Fathom against V8.
        ctx.struct_type_indices
            .insert(te.name.clone(), described_idx);
        ctx.desc_type_indices
            .insert(te.name.clone(), descriptor_idx);
        ctx.struct_fields.insert(te.name.clone(), te.fields.clone());
    }

    // Types with children must be left "open" (`sub`) rather than
    // `sub final` so their subtypes can extend them.
    // Keyed by the declared supertype INDEX, so two same-named-but-distinct
    // types can no longer be conflated into one "has children" answer.
    let mut has_children: std::collections::HashSet<usize> = std::collections::HashSet::new();
    for te in &type_entries {
        if te.parent_index != 0 {
            has_children.insert(te.parent_index as usize);
        }
    }

    // Section header: number of top-level rectypes. One big rec group
    // for the struct pairs (if any) + 3 array singletons + func types +
    // 1 exception type.
    let exception_type_count = 1u32;
    // Stack-switching introduces 2 extra types when used: one func type
    // for the suspend/resume tag, one continuation type wrapping it.
    let ss_extra_types = stack_switching_type_count;
    let rec_group_count = if gc_struct_pairs > 0 { 1u32 } else { 0u32 };
    // A declared indirect signature need not belong to any defined function:
    // mismatching a populated slot must trap at execution, not fail emission.
    let defined_signatures: std::collections::BTreeSet<_> = chunks.iter()
        .filter_map(|chunk| chunk.func_sig.as_ref())
        .map(|(params, results)| format!("{params}->{results}")).collect();
    let extra_call_signatures: std::collections::BTreeSet<_> = chunks.iter()
        .filter(|chunk| !super::callable::source(chunk))
        .flat_map(|chunk| chunk.call_indirect_sigs.values())
        .filter(|signature| !defined_signatures.contains(*signature))
        .cloned().collect();
    let total = rec_group_count
        + array_count
        + func_count
        + block_type_count
        + ss_extra_types
        + exception_type_count
        + 3
        + ctx.generators.len() as u32
        + u32::from(ctx.source_entry.is_some())
        + extra_call_signatures.len() as u32;
    write_leb128_u32(&mut out, total);

    // ── GC struct types: one rec group of (described, descriptor) pairs ──
    if gc_struct_pairs > 0 {
        out.push(GC_REC);
        write_leb128_u32(&mut out, gc_struct_pairs * 2);

        for (i, te) in type_entries.iter().enumerate() {
            let described_idx = (i as u32) * 2;
            let descriptor_idx = (i as u32) * 2 + 1;

            // Opening byte: `sub final` (0x4F) if no subtype extends this,
            // else `sub` (0x50) leaving the type open for extension.
            let described_final = !has_children.contains(&(class_rows[i] + 1));
            let sub_byte = if described_final { CD_SUB_FINAL } else { 0x50 };

            // Described struct subtype. Supertype count = 1 when parent
            // is named and resolvable, else 0. Parent's described typeidx
            // is the supertype link (the descriptor-struct side mirrors
            // by linking to the parent's descriptor).
            out.push(sub_byte);
            if let Some(parent_idx) = ctx.struct_type_by_index(te.parent_index as u32) {
                assert!(
                    parent_idx < described_idx,
                    "WASM superclass must precede its subtype"
                );
                write_leb128_u32(&mut out, 1);
                write_leb128_u32(&mut out, parent_idx);
            } else {
                write_leb128_u32(&mut out, 0);
            }
            out.push(CD_DESCRIPTOR);
            write_leb128_u32(&mut out, descriptor_idx);
            out.push(GC_STRUCT);
            write_leb128_u32(&mut out, te.fields.len() as u32);
            for _ in &te.fields {
                out.push(TYPE_EXTERNREF);
                out.push(GC_MUT);
            }

            // Descriptor struct subtype — same supertype story but keyed
            // off the parent's descriptor index.
            out.push(sub_byte);
            // Each entry emits a described/descriptor PAIR, so the parent's
            // descriptor sits one index past its described type.
            if let Some(parent_idx) = ctx.struct_type_by_index(te.parent_index as u32) {
                write_leb128_u32(&mut out, 1);
                write_leb128_u32(&mut out, parent_idx + 1);
            } else {
                write_leb128_u32(&mut out, 0);
            }
            out.push(CD_DESCRIBES);
            write_leb128_u32(&mut out, described_idx);
            out.push(GC_STRUCT);
            // TWO externref slots before the funcrefs, matching the compiler's
            // row (`append_descriptor_type_rows`: `__desc_proto`,
            // `__desc_props`, then methods in table order).
            //
            // ⛔ This emitted `1 + methods` and left `__desc_props` out
            // entirely, so the two halves disagreed on the layout of the same
            // struct: method `k` sat at compiler index `2+k` and writer index
            // `1+k`. Latent only because nothing reads a descriptor yet — it
            // detonates on the FIRST reader, which is what a field-0/vtable
            // writer is. Field 1's CONTENT is Cairn's; the SLOT has to exist
            // here or neither half can be written.
            let desc_field_count = 2 + te.methods.len();
            write_leb128_u32(&mut out, desc_field_count as u32);
            // field 0 — JS prototype; field 1 — property metadata. BOTH
            // MUTABLE, and the mutability is what makes them fillable at all.
            //
            // Their values are runtime objects, so they cannot appear in the
            // constant init expression that allocates the singleton, and every
            // descriptor field being `GC_IMMUT` meant there was no legal
            // instruction that could ever write them afterwards — the
            // descriptor was sealed empty. `struct.set` needs the field to be
            // mutable.
            //
            // It also sidesteps the ordering problem a `struct.new` runs into:
            // the descriptor's field COUNT depends on the merged method list,
            // which is only final after all emission, whereas `struct.set` of
            // field 0 needs no count.
            //
            // ⚠ Mutability here is OURS to choose — the proposal makes
            // descriptors ordinary structs and does not require immutable
            // fields. The identity guarantee comes from WHAT is stored (the
            // very object `C.prototype` is), not from the field being
            // read-only. Sub/super descriptor field mutability must MATCH for
            // the prefix rule, and it does: every class gets the same shape.
            for _ in 0..2 {
                out.push(TYPE_EXTERNREF);
                out.push(GC_MUT);
            }
            // The vtable stays IMMUTABLE. Fields 0/1 had to become mutable
            // because their values are runtime objects and so cannot appear in
            // a constant init expression — but a method's `ref.func` IS a
            // constant instruction, so the whole vtable can be supplied to the
            // `struct.new` that allocates the singleton and never needs to be
            // written again. Immutable is the honest declaration for a vtable
            // and it costs nothing here.
            for _ in &te.methods {
                out.push(TYPE_FUNCREF);
                out.push(GC_IMMUT);
            }

            // ⛔ The vtable PLAN is built here, in the same iteration that just
            // wrote `desc_field_count`, because `struct.new` carries no count
            // immediate — the operand list it is given IS the field count. Any
            // other assembly point can drift from this one, and the failure is
            // an invalid module rather than a subtle wrong answer.
            //
            // Function index = `import_count + chunk_index`, the same mapping
            // `encode_element_section` uses to fill the funcref table. That
            // segment is also what puts every chunk function in the module's
            // declared-reference set, so `ref.func` on any of them is valid in
            // a constant expression.
            ctx.desc_vtable.push(
                te.methods
                    .iter()
                    .map(|(_, chunk_idx)| {
                        (*chunk_idx < chunks.len()).then(|| (import_count + *chunk_idx) as u32)
                    })
                    .collect(),
            );
        }
    }

    // ── Array types ──
    // Three flavours declared in the order the context recorded above:
    //   array_type_idx        → (array (mut externref)) — dynamic Value arrays
    //   string_array_type_idx → (array (mut i16))       — UTF-16 strings
    //   byte_array_type_idx   → (array (mut i8))        — byte / TypedArray backing
    // Packed types (i8 / i16) are only valid as array/struct field storage
    // types, not as top-level value types — they're emitted with the
    // `PACKED_*` byte tags per the GC proposal.
    out.push(GC_ARRAY);
    out.push(TYPE_EXTERNREF);
    out.push(GC_MUT);
    out.push(GC_ARRAY);
    out.push(PACKED_I16);
    out.push(GC_MUT);
    out.push(GC_ARRAY);
    out.push(PACKED_I8);
    out.push(GC_MUT);

    // Standard native-only Cell, Captures and Callable. Original module type IDs retain their existing mapping.
    out.extend([0x5f, 1, TYPE_EXTERNREF, 1]);
    out.extend([0x5e, 0x64]);
    write_leb128_i64(&mut out, ctx.callables.cell as i64);
    out.push(0);
    out.extend([0x5f, 6, TYPE_FUNCREF, 0, 0x63]);
    write_leb128_i64(&mut out, ctx.callables.captures as i64);
    out.extend([0, TYPE_I32, 0, TYPE_EXTERNREF, 0, TYPE_I32, 0, TYPE_I32, 0]);

    // ── Function types ──
    // Host imports use externref ABI types inferred by arity. Proposal
    // builtins own their exact signatures in their module-specific tables
    // below, including typed wasm:js-number and wasm:js-string helpers.

    // Import function types — per-import typed signatures. Untyped host
    // imports retain the call arity collected for this declaration.
    let host_import_count = host_imports.len();
    for i in 0..host_import_count {
        out.push(TYPE_FUNC);
        // ⛔ ARITY IS NOT A SIGNATURE. A proposal builtin reached through the
        // HOST import table is the same function as one reached through the
        // runtime table — `wasm:js-string.test` returns i32 either way — but
        // this loop used to type every host import `(externref…) -> externref`
        // purely from a bytecode arity scan. A `test` result then fed an
        // `if (result i32)` as an externref and V8 refused the module. Ask the
        // owning proposal module for the real signature; the scan is only the
        // fallback for genuinely untyped host functions.
        let imp = &host_imports[i].import;
        let (module, name) = (imp.module.as_str(), imp.name.as_str());
        let sig_start = out.len();
        if module == crate::writer::builtins::js_array_builtins::MODULE && name == "unshift" {
            crate::writer::builtins::js_array_builtins::write_unshift_signature(
                &mut out,
                host_imports[i].arity.unwrap_or(0),
            );
            record_result_signature(&mut ctx, i as u32, &out, sig_start);
            record_raw_params(&mut ctx, i as u32, &out, sig_start);
        } else if module == crate::writer::builtins::js_array_builtins::MODULE && name == "splice" {
            crate::writer::builtins::js_array_builtins::write_splice_signature(
                &mut out,
                host_imports[i].arity.unwrap_or(0),
            );
            record_result_signature(&mut ctx, i as u32, &out, sig_start);
            record_raw_params(&mut ctx, i as u32, &out, sig_start);
        } else if host_imports[i].arity.is_some_and(|arity| {
            crate::writer::builtins::web_builtins::write_overloaded_signature(
                &mut out, module, name, arity,
            )
        }) || write_proposal_signature(&mut out, module, name, ctx.string_array_type_idx) {
            record_result_signature(&mut ctx, i as u32, &out, sig_start);
            record_raw_params(&mut ctx, i as u32, &out, sig_start);
        } else {
            let argc = host_imports[i].arity.unwrap_or(0);
            write_leb128_u32(&mut out, argc as u32);
            for _ in 0..argc {
                out.push(TYPE_EXTERNREF);
            }
            write_leb128_u32(&mut out, 1);
            out.push(TYPE_EXTERNREF);
            ctx.func_result_arity.insert(i as u32, 1);
            ctx.import_param_funcs
                .insert(i as u32, vec![TYPE_EXTERNREF; argc as usize]);
        }
    }
    // Runtime imports: each proposal module owns the signatures for
    // its own (module, name) pairs. Query them in order, falling back
    // to `(… ) -> externref` for anything unrecognised.
    for (rt_i, &(module, name)) in rt_imports.iter().enumerate() {
        let func_idx = (host_import_count + rt_i) as u32;
        out.push(TYPE_FUNC);
        let sig_start = out.len();
        if write_proposal_signature(&mut out, module, name, ctx.string_array_type_idx) {
            record_result_signature(&mut ctx, func_idx, &out, sig_start);
            record_raw_params(&mut ctx, func_idx, &out, sig_start);
        } else {
            // ⛔ THE DEFAULT IS `() -> externref` — ZERO PARAMETERS. Anything
            // reaching it that is really CALLED with arguments would declare a
            // signature it does not have. Every module whose imports the writer
            // itself emits owns a signature table above; this is only for pairs
            // no proposal claims.
            write_leb128_u32(&mut out, 0);
            write_leb128_u32(&mut out, 1);
            out.push(TYPE_EXTERNREF);
            ctx.func_result_arity.insert(func_idx, 1);
            ctx.import_param_funcs.insert(func_idx, Vec::new());
        }
    }

    for (i, chunk) in chunks.iter().enumerate() {
        let type_idx = ctx.func_type_base + import_count as u32 + i as u32;
        out.push(TYPE_FUNC);
        // ⛔ THE DECLARED SIGNATURE WAS BEING DISCARDED HERE. This loop emitted
        // `externref` for every parameter and result, from `chunk.arity` — a
        // COUNT — while `chunk.func_sig` sat beside it holding the declared
        // value types. So `(func (param i32) (param i32) (result i32))` reached
        // the binary as `(param externref externref) (result externref)` with a
        // body that then ran `i32.add` on them: not a missing check, a module
        // that is genuinely ill-typed as emitted. Confirmed at the bytes by
        // Vesper — the declared pattern `60 02 7F 7C 01 7F` appears ZERO times
        // in a module that declares it, while the all-externref pattern appears
        // nine.
        //
        // `func_sig` is `None` for every chunk that is not a wast function, and
        // those keep the externref shape they had — a dynamic language has no
        // declared wasm signature to emit.
        let declared: Option<(Vec<String>, Vec<String>)> = chunk.func_sig.as_ref().map(|(p, r)| {
            let split = |s: &String| {
                s.split(',')
                    .map(str::trim)
                    .filter(|t| !t.is_empty())
                    .map(str::to_string)
                    .collect::<Vec<_>>()
            };
            (split(p), split(r))
        });
        let signature = match &declared {
            Some((p, r)) => format!("{}->{}", p.join(","), r.join(",")),
            None => String::new(),
        };
        match &declared {
            Some((params, results)) => {
                write_leb128_u32(&mut out, params.len() as u32);
                for t in params {
                    out.push(val_type_byte(t));
                }
                write_leb128_u32(&mut out, results.len() as u32);
                for t in results {
                    out.push(val_type_byte(t));
                }
                ctx.func_result_arity
                    .insert(import_count as u32 + i as u32, results.len() as u8);
            }
            None => {
                // WASM convention: arity params (slot 0 = first arg, no reserved callee slot).
                let param_count = chunk.arity as u32 + 1;
                write_leb128_u32(&mut out, param_count);
                for _ in 0..param_count {
                    out.push(TYPE_EXTERNREF);
                }
                // Multi-value proposal: chunks may return more than one externref.
                let result_count = (chunk.result_arity as u32).max(1);
                write_leb128_u32(&mut out, result_count);
                for _ in 0..result_count {
                    out.push(TYPE_EXTERNREF);
                }
                ctx.func_result_arity
                    .insert(import_count as u32 + i as u32, result_count as u8);
            }
        }
        // Record first type index seen for each arity (for call_ref/call_indirect dispatch)
        ctx.func_type_by_arity
            .entry(chunk.arity)
            .or_insert(type_idx);
        // ⛔ AND BY SIGNATURE, which is what `call_indirect` needs. Keying only
        // by arity is `.or_insert` — first type seen per arity wins — so with
        // real functypes an argc-1 call site would be handed whichever argc-1
        // type came first, and `call_indirect`'s immediate must EQUAL the
        // callee's functype or a conforming engine traps.
        if !signature.is_empty() {
            ctx.func_type_by_signature
                .entry(signature)
                .or_insert(type_idx);
        }
    }

    // Block types: one `externref^M -> externref^N` per distinct
    // (params, results) pair found in the pre-scan. Index is recorded in
    // `ctx.block_type_by_results` for the code emitter to look up when
    // writing an s33 typeidx blocktype.
    let block_type_base = ctx.func_type_base + import_count as u32 + chunks.len() as u32;
    let mut block_type_offset = 0u32;
    for &(params, results) in block_result_counts.iter() {
        let tidx = block_type_base + block_type_offset;
        block_type_offset += 1;
        ctx.block_type_by_results.insert((params, results), tidx);
        out.push(TYPE_FUNC);
        write_leb128_u32(&mut out, params as u32);
        for _ in 0..params {
            out.push(TYPE_EXTERNREF);
        }
        write_leb128_u32(&mut out, results as u32);
        for _ in 0..results {
            out.push(TYPE_EXTERNREF);
        }
    }
    for &params in block_i32_param_counts.iter() {
        let tidx = block_type_base + block_type_offset;
        block_type_offset += 1;
        ctx.block_i32_type_by_params.insert(params, tidx);
        out.push(TYPE_FUNC);
        write_leb128_u32(&mut out, params as u32);
        for _ in 0..params {
            out.push(TYPE_EXTERNREF);
        }
        write_leb128_u32(&mut out, 1);
        out.push(TYPE_I32);
    }

    for &(argc, results) in &source_call_counts {
        let tidx = block_type_base + block_type_offset;
        block_type_offset += 1;
        ctx.callables.calls.insert((argc, results), tidx);
        out.push(TYPE_FUNC);
        write_leb128_u32(&mut out, argc as u32 + 1);
        for _ in 0..=argc {
            out.push(TYPE_EXTERNREF);
        }
        write_leb128_u32(&mut out, results as u32);
        for _ in 0..results {
            out.push(TYPE_EXTERNREF);
        }
    }

    // Stack-switching: suspend/resume tag func type + continuation type.
    // A suspend yields an externref and resumes with an externref, so the
    // tag's signature is `(func (param externref) (result externref))`.
    // The continuation type `(cont $ft)` wraps that func type.
    if uses_stack_switching {
        // (func (param externref) (result externref))
        out.push(TYPE_FUNC);
        write_leb128_u32(&mut out, 1); // 1 param
        out.push(TYPE_EXTERNREF);
        write_leb128_u32(&mut out, 1); // 1 result
        out.push(TYPE_EXTERNREF);
        // (cont $ft) — prefix + funcidx
        out.push(crate::writer::proposals::stack_switching::CONT_TYPE_PREFIX);
        // The continuation's function type is an s33, unlike instruction
        // type-index immediates. Unsigned LEB interprets indices 64..127 as
        // negative values when the decoder reads the required signed encoding.
        write_leb128_i64(&mut out, ctx.suspend_tag_type_idx as i64);
        // Native continuation carrier: resumable reference + observable owner.
        out.push(0x5f); // struct
        write_leb128_u32(&mut out, 3);
        out.push(0x63);
        write_leb128_i64(&mut out, ctx.continuation_type_idx as i64);
        out.push(1);
        out.push(TYPE_EXTERNREF);
        out.push(1);
        out.push(0x70); // originating raw driver (funcref)
        out.push(0); // immutable
        // Suspension landing: tag payload followed by the new continuation.
        out.push(TYPE_FUNC);
        write_leb128_u32(&mut out, 0);
        write_leb128_u32(&mut out, 2);
        out.push(TYPE_EXTERNREF);
        out.push(0x63);
        write_leb128_i64(&mut out, ctx.continuation_type_idx as i64);
    }

    if native_next {
        out.push(TYPE_FUNC);
        write_leb128_u32(&mut out, 2);
        out.push(TYPE_EXTERNREF);
        out.push(0x63);
        write_leb128_i64(&mut out, ctx.continuation_type_idx as i64);
        write_leb128_u32(&mut out, 2);
        out.push(TYPE_EXTERNREF);
        out.push(0x63);
        write_leb128_i64(&mut out, ctx.continuation_type_idx as i64);
    }

    if native_next {
        out.push(TYPE_FUNC);write_leb128_u32(&mut out,2);out.extend([TYPE_EXTERNREF;2]);
        write_leb128_u32(&mut out,2);out.push(TYPE_EXTERNREF);out.push(0x63);write_leb128_i64(&mut out,ctx.continuation_type_idx as i64);
    }

    for tag in &continuation_tags {
        out.push(TYPE_FUNC);
        write_leb128_u32(&mut out, 1);
        out.push(continuation_tag_valtype(&tag.yield_type));
        write_leb128_u32(&mut out, 1);
        out.push(continuation_tag_valtype(&tag.resume_type));
    }

    // Exception tag type — `(externref) -> ()` per the exception-handling
    // proposal. The tag section references this type index.
    out.push(TYPE_FUNC);
    write_leb128_u32(&mut out, 1); // 1 param
    out.push(TYPE_EXTERNREF);
    write_leb128_u32(&mut out, 0); // 0 results

    for e in &ctx.generators {
        out.push(crate::writer::proposals::stack_switching::CONT_TYPE_PREFIX);
        write_leb128_i64(&mut out, e.body_type as i64);
    }
    if ctx.source_entry.is_some() {
        out.extend_from_slice(&[TYPE_FUNC, 0, 0]);
    }
    let extra_base = ctx.exception_type_idx + 1 + ctx.generators.len() as u32
        + u32::from(ctx.source_entry.is_some());
    for (offset, signature) in extra_call_signatures.into_iter().enumerate() {
        let (params, results) = signature.split_once("->").expect("declared call signature");
        out.push(TYPE_FUNC);
        for list in [params, results] {
            let types: Vec<_> = list.split(',').map(str::trim).filter(|ty| !ty.is_empty()).collect();
            write_leb128_u32(&mut out, types.len() as u32);
            for ty in types { out.push(val_type_byte(ty)); }
        }
        ctx.func_type_by_signature.entry(signature).or_insert(extra_base + offset as u32);
    }
    (out, ctx)
}

#[cfg(test)]
mod superclass_order_tests {
    use super::*;
    use vybe_runtime::{Op, chunk::TypeEntry};

    fn class(name: &str, parent: u16, fields: &[&str]) -> TypeEntry {
        TypeEntry {
            name: name.into(),
            parent_index: parent,
            fields: fields.iter().map(|s| (*s).into()).collect(),
            ..Default::default()
        }
    }

    #[test]
    fn repeated_diagnostic_names_keep_distinct_descriptor_singletons() {
        let mut root = Chunk::new("repeated-type-names");
        root.types = vec![class("Repeated", 0, &[]), class("Repeated", 0, &[])];
        root.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
        root.emit_op(Op::RETURN, 0);
        let chunks = vec![root];
        let (_, mut ctx) = build_type_context(&chunks, &[], 0, &[]);
        assert_eq!(ctx.desc_type_indices.len(), 1);
        assert_eq!(ctx.descriptor_global_count(), 2);
        ctx.desc_global_base = Some(20);
        assert_eq!(ctx.desc_global_by_index(1), Some(20));
        assert_eq!(ctx.desc_global_by_index(2), Some(21));
        let bytes = crate::write_wasm(&chunks);
        wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
            .validate_all(&bytes)
            .unwrap();
    }
    fn chunks() -> Vec<Chunk> {
        let mut root = Chunk::new("ordered-types");
        let mut child = class("Child", 3, &["base", "leaf"]);
        child.methods = vec![("inherited".into(), 1), ("local".into(), 2)];
        let mut parent = class("Parent", 0, &["base"]);
        parent.methods = vec![("inherited".into(), 1)];
        let mut child_desc = class("#desc Child", 0, &[]);
        child_desc.describes_index = 1;
        let mut parent_desc = class("#desc Parent", 0, &[]);
        parent_desc.describes_index = 3;
        root.types = vec![child, child_desc, parent, parent_desc];
        // Original module type 1 is Child. The allocation must use its new
        // binary type and matching descriptor singleton, without changing ID 1.
        root.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
        root.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
        // This is the compiler allocation producer: the writer supplies the
        // exact descriptor singleton, then externalizes the allocated value.
        // Raw STRUCT_NEW_DESC instead requires its descriptor on the stack.
        root.emit_struct_new(1, 2, 0);
        root.emit_op(Op::DROP, 0);
        root.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
        root.emit_op(Op::RETURN, 0);
        let mut method = Chunk::new("inherited");
        method.emit_ref_null(vybe_runtime::opcode::heaptype::HT_EXTERN, 0);
        method.emit_op(Op::RETURN, 0);
        let mut local = method.clone();
        local.name = "local".into();
        vec![root, method, local]
    }
    #[test]
    fn forward_superclasses_remap_module_ids_descriptors_and_vtables_together() {
        let chunks = chunks();
        let before = format!("{:?}", chunks[0].types);
        let (_, mut ctx) = build_type_context(&chunks, &[], 0, &[]);
        assert_eq!(ctx.struct_type_by_module_index, vec![2, 3, 0, 1]);
        assert_eq!(ctx.struct_type("Child"), Some(2));
        assert_eq!(ctx.struct_type("Parent"), Some(0));
        assert_eq!(ctx.desc_type("Child"), Some(3));
        assert_eq!(ctx.desc_type("Parent"), Some(1));
        ctx.desc_global_base = Some(20);
        assert_eq!(ctx.desc_global_by_index(1), Some(21));
        assert_eq!(ctx.desc_global_by_index(3), Some(20));
        assert_eq!(ctx.desc_global_by_index(2), None);
        assert_eq!(ctx.desc_global_by_index(4), None);
        assert_eq!(ctx.desc_global("Child"), Some(21));
        // Explicit descriptor rows enlarge the module table, not class pairs.
        assert_eq!(ctx.desc_global_for_described(2), Some(21));
        assert_eq!(ctx.desc_global_for_described(3), None);
        assert_eq!(ctx.desc_global_for_described(4), None);
        assert_eq!(ctx.desc_global_for_described(6), None);
        assert_eq!(ctx.desc_vtable, vec![vec![Some(1)], vec![Some(1), Some(2)]]);
        assert_eq!(ctx.field_index("Child", "leaf"), Some(1));
        assert_eq!(format!("{:?}", chunks[0].types), before);
    }
    #[test]
    fn ordered_native_binary_validates_all_superclass_and_descriptor_links() {
        let chunks = chunks();
        let bytes = crate::write_wasm(&chunks);
        wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
            .validate_all(&bytes)
            .unwrap();
        for payload in wasmparser::Parser::new(0).parse_all(&bytes) {
            if let wasmparser::Payload::TypeSection(types) = payload.unwrap() {
                let group = types.into_iter().next().unwrap().unwrap();
                let defs: Vec<_> = group.types().collect();
                assert_eq!(defs.len(), 4);
                let parent = |t: &wasmparser::SubType| {
                    assert!(t.supertype_idxs.len() <= 1);
                    t.supertype_idxs.first().map(|index| match index.unpack() {
                        wasmparser::UnpackedIndex::Module(index) => index,
                        other => panic!("unexpected superclass index {other:?}"),
                    })
                };
                assert_eq!(
                    defs.iter().map(|t| parent(t)).collect::<Vec<_>>(),
                    vec![None, None, Some(0), Some(1)]
                );
                assert!(!defs[0].is_final);
                assert!(!defs[1].is_final);
                assert!(defs[2].is_final);
                assert!(defs[3].is_final);
            }
        }
    }
    #[test]
    fn ordering_is_deterministic_for_reversed_chains_siblings_and_case_distinct_names() {
        let rows = vec![
            class("Foo", 3, &[]),
            class("FOO", 3, &[]),
            class("Middle", 4, &[]),
            class("Base", 0, &[]),
        ];
        assert_eq!(ordered_class_rows(&rows).unwrap(), vec![3, 2, 0, 1]);
        let mut root = Chunk::new("case-types");
        root.types = rows;
        let (_, ctx) = build_type_context(&[root], &[], 0, &[]);
        assert_eq!(ctx.struct_type("Foo"), Some(4));
        assert_eq!(ctx.struct_type("FOO"), Some(6));
    }
    #[test]
    fn invalid_parent_and_descriptor_indices_and_cycles_are_rejected() {
        assert!(
            ordered_class_rows(&[class("Self", 1, &[])])
                .unwrap_err()
                .contains("cycle")
        );
        assert!(
            ordered_class_rows(&[class("A", 2, &[]), class("B", 1, &[])])
                .unwrap_err()
                .contains("cycle")
        );
        assert!(
            ordered_class_rows(&[class("A", 2, &[])])
                .unwrap_err()
                .contains("out-of-range parent")
        );
        let mut desc = class("#desc A", 0, &[]);
        desc.describes_index = 1;
        assert!(
            ordered_class_rows(&[class("A", 2, &[]), desc])
                .unwrap_err()
                .contains("descriptor parent")
        );
        let mut desc = class("#desc missing", 0, &[]);
        desc.describes_index = 2;
        assert!(
            ordered_class_rows(&[desc])
                .unwrap_err()
                .contains("invalid class")
        );
        let mut root = Chunk::new("invalid");
        root.types = vec![class("A", 1, &[])];
        assert!(
            std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| crate::write_wasm(&[root])))
                .is_err()
        );
    }
}

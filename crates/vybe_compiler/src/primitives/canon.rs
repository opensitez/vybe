//! `vybe_ast::canon::CanonDecl` → `vybe_runtime::canon_def::CanonDef`.
//!
//! The front end states what the SOURCE said; the runtime holds what it needs
//! to EXECUTE. This is the conversion between them, and it runs once at
//! emission — the same split, and the same moment, as
//! `globals::normalize_global_table`.
//!
//! Nothing here defaults an absent immediate into a plausible value. The one
//! default that IS applied is `string-encoding`, because `Binary.md` states it:
//! "When there is no `string-encoding` present, the default value is `utf8`.
//! For all other values, the default value is 'none'."

use std::collections::HashMap;
use std::sync::Arc;
use vybe_ast::canon::{CanonCallee, CanonDecl, ResourceRep, TypeDecl, ValSpec};
use vybe_runtime::canon_def::{CanonDef, CanonFuncType, CanonOpts, CanonSection, StringEncoding};
use vybe_runtime::chunk::Chunk;
use vybe_runtime::component::ValType;

const EMITTER_ROW: &str = "@emitter:";

/// Retain a helper's canonical type until module-owned indices are assigned.
/// The symbolic suffix is compiler-only and is resolved before chunks escape.
pub fn add_typed_stream_import(
    chunk: &mut Chunk,
    builtin: &str,
    element: ValType,
    opts: CanonOpts,
) -> u16 {
    let ty = ValType::Stream(Box::new(element));
    let types = Arc::make_mut(&mut chunk.canon_valtypes);
    let index = match types.iter().position(|entry| entry.as_ref() == Some(&ty)) {
        Some(index) => index,
        None => { types.push(Some(ty)); types.len() - 1 }
    };
    let definition = CanonDef::with_type_and_opts(
        u32::try_from(index).expect("canonical type index fits u32"), opts);
    let definitions = Arc::make_mut(&mut chunk.canon_section);
    let row = match definitions.iter().position(|entry| entry == &definition) {
        Some(index) => index,
        None => { definitions.push(definition); definitions.len() - 1 }
    };
    chunk.add_import("canon", &format!("{builtin}{EMITTER_ROW}{row}"))
}

/// Relocate helper-owned stream definitions into the module's existing spaces.
/// Source canon indices keep their meaning, including when helpers share a chunk
/// with source canonical imports. Import positions and call instructions do not move.
pub fn resolve_emitter_imports(
    chunks: &mut [Chunk],
    definitions: &mut Vec<CanonDef>,
    functypes: &mut Vec<Option<CanonFuncType>>,
    valtypes: &mut Vec<Option<ValType>>,
) -> Result<(), String> {
    for chunk in chunks {
        if !chunk.imports.iter().any(|import|
            import.module == "canon" && import.name.contains(EMITTER_ROW))
        {
            continue;
        }
        let local_definitions = chunk.canon_section.clone();
        let local_types = chunk.canon_valtypes.clone();
        for import in &mut chunk.imports {
            if import.module != "canon" { continue; }
            let Some((builtin, suffix)) = import.name.split_once(EMITTER_ROW) else { continue; };
            let row = suffix.parse::<usize>().map_err(|_| "invalid emitter canonical row")?;
            let local = local_definitions.get(row).ok_or("missing emitter canonical row")?;
            let ty = local.ty.and_then(|ty| local_types.get(ty as usize))
                .and_then(Option::as_ref).ok_or("missing emitter canonical stream type")?;
            if !matches!(ty, ValType::Stream(_))
                || *local != CanonDef::with_type_and_opts(local.ty.unwrap(), local.opts.clone())
            {
                return Err("emitter stream definition carries unsupported metadata".into());
            }
            let index = match valtypes.iter().position(|entry| entry.as_ref() == Some(ty)) {
                Some(index) => index,
                None => {
                    let index = valtypes.len().max(functypes.len());
                    valtypes.resize(index, None);
                    functypes.resize(index + 1, None);
                    valtypes.push(Some(ty.clone()));
                    index
                }
            };
            let definition = CanonDef::with_type_and_opts(
                u32::try_from(index).map_err(|_| "canonical type space exhausted")?,
                local.opts.clone());
            let row = match definitions.iter().position(|entry| entry == &definition) {
                Some(row) => row,
                None => { definitions.push(definition); definitions.len() - 1 }
            };
            let row = u32::try_from(row).map_err(|_| "canonical definition space exhausted")?;
            import.name = format!("{builtin}@{row}");
        }
    }
    Ok(())
}

/// The core value type of a `context.get` / `context.set` slot.
fn context_valtype(spelling: &str) -> Result<ValType, String> {
    match spelling {
        "i32" => Ok(ValType::I32),
        "i64" => Ok(ValType::I64),
        other => Err(format!(
            "canon context.*: `{other}` is not an i32 or i64 core value type"
        )),
    }
}

fn string_encoding(spelling: &Option<String>) -> Result<StringEncoding, String> {
    match spelling.as_deref() {
        // Binary.md: absent means utf8.
        None | Some("utf8") => Ok(StringEncoding::Utf8),
        Some("utf16") => Ok(StringEncoding::Utf16),
        Some("latin1+utf16") => Ok(StringEncoding::Latin1Utf16),
        Some(other) => Err(format!("canon: `{other}` is not a string-encoding")),
    }
}

/// One declared row → one runtime row.
pub fn lower_decl(d: &CanonDecl) -> Result<CanonDef, String> {
    let is_worker = matches!(d.builtin.as_str(),
        "thread.new-indirect" | "thread.spawn-indirect" | "thread.spawn-ref");
    if is_worker != d.worker_param.is_some() {
        return Err(format!("canon {}: missing or unexpected core worker signature", d.builtin));
    }
    let is_context = matches!(d.builtin.as_str(), "context.get" | "context.set");
    if is_context && d.context.is_none() {
        return Err(format!(
            "canon {}: missing core value type and slot immediate",
            d.builtin
        ));
    }
    if !is_context && d.context.is_some() {
        return Err(format!(
            "canon {}: unexpected context type and slot immediate",
            d.builtin
        ));
    }
    Ok(CanonDef {
        callee: match &d.callee {
            Some(CanonCallee::Core(i)) => Some(vybe_runtime::canon_def::CalleeRef::Core(*i)),
            Some(CanonCallee::Component(i)) => {
                Some(vybe_runtime::canon_def::CalleeRef::Component(*i))
            }
            // Left UNRESOLVED on purpose. A `CoreExport` needs the chunk the
            // method compiled to, and no chunk exists yet at this point —
            // `resolve_core_export_callees` fills it in after emission. `None`
            // here is not a default: if that pass never runs, `require_callee`
            // traps with "carries no $callee immediate" rather than calling
            // something arbitrary.
            Some(CanonCallee::CoreExport { .. }) | None => None,
        },
        resource_dtor: None,
        resource_rep: None,
        ty: d.ty,
        functype: d.functype,
        results: d.results.clone(),
        context: match &d.context {
            Some((v, i)) => {
                if *i >= 2 {
                    return Err(format!("canon context.*: slot {i} must be less than 2"));
                }
                Some((context_valtype(v)?, *i))
            }
            None => None,
        },
        table: d.table,
        worker_param: match d.worker_param {
            Some(vybe_ast::canon::ComponentCoreValType::I32) => Some(vybe_runtime::canon_flat::CoreType::I32),
            Some(vybe_ast::canon::ComponentCoreValType::I64) => Some(vybe_runtime::canon_flat::CoreType::I64),
            Some(_) => return Err("canon thread worker must take i32 or i64".into()),
            None => None,
        },
        opts: CanonOpts {
            string_encoding: string_encoding(&d.opts.string_encoding)?,
            memory: d.opts.memory,
            realloc: match &d.opts.realloc {
                Some(vybe_ast::canon::CanonCoreFuncRef::Index(i)) => Some(*i),
                _ => None,
            },
            post_return: match &d.opts.post_return {
                Some(vybe_ast::canon::CanonCoreFuncRef::Index(i)) => Some(*i),
                _ => None,
            },
            callback: match &d.opts.callback {
                Some(vybe_ast::canon::CanonCoreFuncRef::Index(i)) => Some(*i),
                _ => None,
            },
            is_async: d.opts.is_async,
        },
        cancellable: d.cancellable,
        is_async: d.is_async,
        shared: d.shared,
    })
}

/// Lower a whole declared section, preserving canonidx order.
pub fn lower_section(decls: &[CanonDecl]) -> Result<Vec<CanonDef>, String> {
    let lowered: Vec<CanonDef> = decls
        .iter()
        .map(|d| {
            lower_decl(d).map_err(|e| {
                // Quote the row as the source spelled it — a canonidx alone
                // names nothing a reader can find.
                format!("canon {}: {e}", d.builtin)
            })
        })
        .collect::<Result<_, _>>()?;
    let mut context_type = None;
    for (index, row) in lowered.iter().enumerate() {
        if let Some((ty, _)) = &row.context {
            match &context_type {
                Some(first) if *first != ty => {
                    return Err(format!(
                        "canon row {index}: context.get/set must use the same core value type"
                    ));
                }
                None => context_type = Some(ty),
                _ => {}
            }
        }
    }
    Ok(lowered)
}

pub struct CoreExportCallee {
    row: usize,
    builtin: String,
    class: String,
    method: String,
    target: CoreExportTarget,
}

#[derive(Clone, Copy)]
enum CoreExportTarget {
    Callee,
    ResourceDtor,
    Realloc,
    PostReturn,
    Callback,
}

pub fn collect_core_export_callees(decls: &[CanonDecl]) -> Vec<CoreExportCallee> {
    use vybe_ast::canon::CanonCoreFuncRef;
    let mut pending = Vec::new();
    for (row, decl) in decls.iter().enumerate() {
        let mut add = |class: &str, method: &str, target| pending.push(CoreExportCallee {
            row,
            builtin: decl.builtin.clone(),
            class: class.into(),
            method: method.into(),
            target,
        });
        if let Some(CanonCallee::CoreExport { class, method }) = &decl.callee {
            add(class, method, CoreExportTarget::Callee);
        }
        for (reference, target) in [
            (&decl.opts.realloc, CoreExportTarget::Realloc),
            (&decl.opts.post_return, CoreExportTarget::PostReturn),
            (&decl.opts.callback, CoreExportTarget::Callback),
        ] {
            if let Some(CanonCoreFuncRef::CoreExport { class, method }) = reference {
                add(class, method, target);
            }
        }
    }
    pending
}

pub fn attach_resource_metadata(
    decls: &[CanonDecl],
    types: &[TypeDecl],
    defs: &mut [CanonDef],
) -> Result<Vec<CoreExportCallee>, String> {
    let mut pending = Vec::new();
    for (row, decl) in decls.iter().enumerate() {
        if !matches!(
            decl.builtin.as_str(),
            "resource.new" | "resource.rep" | "resource.drop"
        ) {
            continue;
        }
        let Some(ty) = decl.ty else {
            return Err(format!(
                "canon {}: missing resource type immediate",
                decl.builtin
            ));
        };
        let (rep, dtor) = match types.get(ty as usize) {
            Some(TypeDecl::Resource { rep, dtor, .. }) => (rep, dtor),
            Some(_) => {
                return Err(format!(
                    "canon {}: typeidx {ty} is not a resource",
                    decl.builtin
                ));
            }
            None if types.is_empty() => continue,
            None => {
                return Err(format!(
                    "canon {}: typeidx {ty} is out of range",
                    decl.builtin
                ));
            }
        };
        defs[row].resource_rep = Some(match rep {
            ResourceRep::I32 => vybe_runtime::canon_flat::CoreType::I32,
            ResourceRep::I64 => vybe_runtime::canon_flat::CoreType::I64,
        });
        if decl.builtin != "resource.drop" {
            continue;
        }
        match dtor {
            Some(CanonCallee::Core(i)) => {
                defs[row].resource_dtor = Some(vybe_runtime::canon_def::CalleeRef::Core(*i));
            }
            Some(CanonCallee::CoreExport { class, method }) => pending.push(CoreExportCallee {
                row,
                builtin: "resource.drop destructor".into(),
                class: class.clone(),
                method: method.clone(),
                target: CoreExportTarget::ResourceDtor,
            }),
            Some(CanonCallee::Component(_)) => {
                return Err("canon resource.drop: destructor must name a core function".into());
            }
            None => {}
        }
    }
    Ok(pending)
}

/// Publish the section to every chunk, exactly as `normalize_global_table`
/// publishes the global index space.
///
/// Every chunk, not just chunk 0: a chunk that carries a canonidx has to be
/// able to say what that index means, and threading the table through every
/// caller is what dragged `cli.rs` in when the global table was parked on
/// chunk 0.
pub fn install_canon_section(chunks: &mut [Chunk], section: Vec<CanonDef>) {
    if section.is_empty() {
        return;
    }
    let shared: CanonSection = Arc::new(section);
    for chunk in chunks.iter_mut() {
        chunk.canon_section = shared.clone();
    }
}

/// `vybe_ast::canon::ValSpec` → `vybe_runtime::component::ValType`.
///
/// Refuses every spelling the runtime cannot represent FAITHFULLY, rather than
/// mapping it onto the nearest thing that compiles.
///
/// Narrow integers and both signednesses of 32- and 64-bit integers remain
/// distinct until canonical lifting; they share core Wasm widths, but not
/// component-level numeric values.
struct LowerContext<'a> {
    types: &'a [TypeDecl],
    // ValType owns nested fields, so caching expansions would still deep-clone
    // each use. Track only active references to reject unrepresentable cycles.
    visiting: Vec<bool>,
}

impl<'a> LowerContext<'a> {
    fn new(types: &'a [TypeDecl]) -> Self {
        Self {
            types,
            visiting: Vec::new(),
        }
    }

    fn lower_ref(&mut self, i: u32) -> Result<ValType, String> {
        let index = i as usize;
        let types = self.types;
        let inner = match types.get(index) {
            Some(TypeDecl::Value(inner)) => inner,
            Some(TypeDecl::Resource { .. }) => {
                return Err(format!(
                    "`(type {i})` names a RESOURCE used directly as a value type; a \
                     resource travels as `(own {i})` or `(borrow {i})`, never bare"
                ));
            }
            Some(TypeDecl::Func { .. }) => {
                return Err(format!(
                    "`(type {i})` names a FUNCTION type used as a value type"
                ));
            }
            Some(TypeDecl::Opaque(k)) => {
                return Err(format!(
                    "`(type {i})` names a `{k}`, which is not a value type"
                ));
            }
            None => {
                return Err(format!(
                    "`(type {i})` is not in the component type space (have {})",
                    types.len()
                ));
            }
        };
        if self.visiting.is_empty() {
            self.visiting.resize(types.len(), false);
        }
        if self.visiting[index] {
            return Err(format!("`(type {i})` contains a recursive value type"));
        }
        self.visiting[index] = true;
        let result = lower_valspec_in_context(inner, self);
        self.visiting[index] = false;
        result
    }
}

pub fn lower_valspec(v: &ValSpec, types: &[TypeDecl]) -> Result<ValType, String> {
    lower_valspec_in_context(v, &mut LowerContext::new(types))
}

fn lower_valspec_in_context(v: &ValSpec, ctx: &mut LowerContext<'_>) -> Result<ValType, String> {
    let types = ctx.types;
    Ok(match v {
        ValSpec::Prim(p) => match p.as_str() {
            "bool" => ValType::Bool,
            "s32" => ValType::I32,
            "u32" => ValType::U32,
            "s64" => ValType::I64,
            "u64" => ValType::U64,
            "f64" => ValType::F64,
            "string" => ValType::String,
            // The NARROW widths are their own types, not `I32` with a smaller
            // range: `load` reads exactly 1 or 2 bytes and sign-extends only
            // the signed ones, and the flat lift narrows before widening back
            // into the slot. All four used to refuse for exactly that reason.
            "s8" => ValType::S8,
            "u8" => ValType::U8,
            "s16" => ValType::S16,
            "u16" => ValType::U16,
            // A real 32-bit float — four bytes in memory, a core `f32` flat.
            "f32" => ValType::F32,
            // A Unicode SCALAR value. `convert_i32_to_char` traps on the
            // surrogate range and on anything at or above 0x110000, which is
            // the entire difference between this and `u32`.
            "char" => ValType::Char,
            // 📝 The value lives in the instance handle table; the type is
            // only the i32 handle's static type.
            "error-context" => ValType::ErrorContext,
            other => return Err(format!("`{other}` is not a primitive value type")),
        },
        ValSpec::List(t) => ValType::List(Box::new(lower_valspec_in_context(t, ctx)?)),
        ValSpec::Option(t) => ValType::Option(Box::new(lower_valspec_in_context(t, ctx)?)),
        ValSpec::Result(ok, err) => ValType::Result(
            match ok {
                Some(t) => Some(Box::new(lower_valspec_in_context(t, ctx)?)),
                None => None,
            },
            match err {
                Some(t) => Some(Box::new(lower_valspec_in_context(t, ctx)?)),
                None => None,
            },
        ),
        ValSpec::Record(fields) => {
            let mut out = Vec::with_capacity(fields.len());
            for (n, t) in fields {
                out.push((n.clone(), lower_valspec_in_context(t, ctx)?));
            }
            ValType::Record(out)
        }
        ValSpec::Variant(cases) => {
            let mut out = Vec::with_capacity(cases.len());
            for (n, t) in cases {
                out.push((
                    n.clone(),
                    match t {
                        Some(t) => Some(lower_valspec_in_context(t, ctx)?),
                        None => None,
                    },
                ));
            }
            ValType::Variant(out)
        }
        // `stream`/`future` with an ABSENT element type is a known deviation
        // already recorded in cmplan.md: the runtime's `Stream(Box<ValType>)`
        // has nowhere to put "no element", and a stand-in would feed
        // `elem_size` and get the layout wrong with nothing reporting it.
        ValSpec::Stream(Some(t)) => ValType::Stream(Box::new(lower_valspec_in_context(t, ctx)?)),
        ValSpec::Future(Some(t)) => ValType::Future(Box::new(lower_valspec_in_context(t, ctx)?)),
        ValSpec::Stream(None) => {
            return Err(
                "`(stream)` with no element type — `ValType::Stream` has nowhere to \
                 record the absence, and a stand-in element would feed `elem_size`"
                    .to_string(),
            );
        }
        ValSpec::Future(None) => {
            return Err(
                "`(future)` with no element type — `ValType::Future` has nowhere to \
                 record the absence, and a stand-in element would feed `elem_size`"
                    .to_string(),
            );
        }
        // ── Specialized types: DESPECIALIZED, not represented ────────────────
        //
        // `CanonicalABI.md:2175` defines `despecialize()`, and every layout,
        // lift and lower rule in that document matches on `despecialize(t)`
        // rather than on `t`. So the spec's own canonical ABI for these three
        // IS the expansion — expanding here is not a shortcut around a missing
        // `ValType` variant, it is what the ABI says they mean.
        //
        //   tuple<ts>   ↦ record with field names "0", "1", …
        //   enum<ls>    ↦ variant, every case payload-less
        //   map<k,v> 🗺️ ↦ list<tuple<k,v>> ↦ list<record{"0":k,"1":v}>
        //
        // ⛔ `string` and `flags` are DELIBERATELY absent from that list, and
        // the paragraph under it says why: they have canonical ABI
        // representations distinct from their expansions. `string` is already
        // a `ValType`; `flags` is genuinely missing and stays refused below.
        // Expanding `flags` to a record of `bool` would lay it out one byte
        // per flag where the ABI bit-packs it into a single integer, which is
        // a wrong layout nothing would report.
        ValSpec::Tuple(ts) => {
            let mut out = Vec::with_capacity(ts.len());
            for (i, t) in ts.iter().enumerate() {
                out.push((i.to_string(), lower_valspec_in_context(t, ctx)?));
            }
            ValType::Record(out)
        }
        ValSpec::Enum(labels) => {
            ValType::Variant(labels.iter().map(|l| (l.clone(), None)).collect())
        }
        ValSpec::Map(k, v) => ValType::List(Box::new(ValType::Record(vec![
            ("0".to_string(), lower_valspec_in_context(k, ctx)?),
            ("1".to_string(), lower_valspec_in_context(v, ctx)?),
        ]))),
        // The one specialized type `despecialize()` does NOT expand
        // (CanonicalABI.md:2185): it BIT-PACKS into one integer of 1, 2 or 4
        // bytes rather than taking a byte per flag, so it has a `ValType` of
        // its own rather than becoming a record of `bool`.
        ValSpec::Flags(labels) => {
            // ⛔ `CanonicalABI.md:2294` asserts `0 < n <= 32`, and it is not a
            // stylistic bound: flags flatten to ONE core `i32`, so a 33rd label
            // has no bit to live in and would be dropped silently. An empty
            // list is refused for the same reason the spec's assertion
            // excludes it — zero labels have no width to pack into.
            if labels.is_empty() || labels.len() > 32 {
                return Err(format!(
                    "`flags` with {} labels — the canonical ABI packs them into ONE i32, \
                     so the spec asserts `0 < n <= 32` (CanonicalABI.md:2294). Beyond 32 \
                     there is no bit to carry the flag and it would be dropped silently",
                    labels.len()
                ));
            }
            ValType::Flags(labels.clone())
        }
        // 🔧 A fixed-length list is a DIFFERENT shape from `List`, not a
        // `List` with a constraint: its elements are inline, so it has no
        // (ptr, len) pair and `store` never calls `realloc`.
        ValSpec::ListFixed(elem, n) => {
            // ⛔ Zero elements would flatten to NOTHING and occupy no bytes,
            // so a `list<T, 0>` parameter would silently vanish from the
            // signature — the caller would pass one fewer core value than the
            // callee expects and every later argument would shift.
            if *n == 0 {
                return Err(
                    "`(list T 0)` \u{1F527} — a zero-length fixed list flattens to no core \
                     values at all, so it would disappear from the signature and shift \
                     every argument after it"
                        .to_string(),
                );
            }
            ValType::ListFixed(Box::new(lower_valspec_in_context(elem, ctx)?), *n)
        }
        // ⛔ The runtime names a resource by STRING and the source names it by
        // INDEX. `TypeDecl::Resource { name, .. }` is the only record connecting the
        // two, and it is filled from the `(type $r …)` binder.
        ValSpec::Own(i) => ValType::Own(resource_name(*i, types, "own")?),
        ValSpec::Borrow(i) => ValType::Borrow(resource_name(*i, types, "borrow")?),
        // A typeidx used AS a value type — `(type $t u32)` then `(param "a" $t)`.
        // It resolves to whatever that index declares, and refuses if the index
        // declares something that is not a value.
        ValSpec::Ref(i) => ctx.lower_ref(*i)?,
    })
}

/// The component's declared type space → the VM's two typeidx-indexed tables.
///
/// Both results are POSITIONALLY ALIGNED with the source type space: entry `n`
/// is source typeidx `n`, and a slot holding the other kind is `None`. Two
/// dense vectors would renumber — a source space of `(type $t u32)` then
/// `(type $ft (func …))` would put the functype at dense index 0 while the
/// source calls it 1, and `canon lift $ft` would then read whatever sat there.
/// That is the `GLOBAL_GET` defect, and alignment is what forecloses it.
///
/// A slot is `None` for the OTHER kind, never a default: `canon lift` naming a
/// value type has to be told it named a value type, not handed an empty
/// signature that lifts zero arguments and looks plausible.
pub fn lower_types(
    types: &[TypeDecl],
) -> Result<(Vec<Option<CanonFuncType>>, Vec<Option<ValType>>), String> {
    let mut functypes = Vec::with_capacity(types.len());
    let mut valtypes = Vec::with_capacity(types.len());
    let mut ctx = LowerContext::new(types);
    for (i, t) in types.iter().enumerate() {
        match t {
            TypeDecl::Func { params, result, is_async } => {
                let mut ps = Vec::with_capacity(params.len());
                for (n, v) in params {
                    ps.push(
                        lower_valspec_in_context(v, &mut ctx)
                            .map_err(|e| format!("component type {i}: param \"{n}\": {e}"))?,
                    );
                }
                let r = match result {
                    Some(v) => Some(
                        lower_valspec_in_context(v, &mut ctx)
                            .map_err(|e| format!("component type {i}: result: {e}"))?,
                    ),
                    None => None,
                };
                functypes.push(Some(CanonFuncType {
                    params: ps,
                    result: r,
                    is_async: *is_async,
                }));
                valtypes.push(None);
            }
            TypeDecl::Value(v) => {
                let lowered = lower_valspec_in_context(v, &mut ctx)
                    .map_err(|e| format!("component type {i}: {e}"))?;
                functypes.push(None);
                valtypes.push(Some(lowered));
            }
            // Occupies its index and nothing more — which is the point: it must
            // still advance the space, or every later typeidx shifts.
            // A resource occupies its index and is neither a function type
            // nor a value type — a handle to it is `own`/`borrow`.
            TypeDecl::Resource { .. } => {
                functypes.push(None);
                valtypes.push(None);
            }
            TypeDecl::Opaque(_) => {
                functypes.push(None);
                valtypes.push(None);
            }
        }
    }
    Ok((functypes, valtypes))
}

/// Publish the type space to every chunk, exactly as `install_canon_section`
/// publishes the canon section.
pub fn install_type_space(
    chunks: &mut [Chunk],
    functypes: Vec<Option<CanonFuncType>>,
    valtypes: Vec<Option<ValType>>,
    funcs: Vec<Option<u32>>,
) {
    if functypes.is_empty() && valtypes.is_empty() && funcs.is_empty() {
        return;
    }
    let fu: Arc<Vec<Option<u32>>> = Arc::new(funcs);
    let f: Arc<Vec<Option<CanonFuncType>>> = Arc::new(functypes);
    let v: Arc<Vec<Option<ValType>>> = Arc::new(valtypes);
    for chunk in chunks.iter_mut() {
        chunk.canon_functypes = f.clone();
        chunk.canon_valtypes = v.clone();
        chunk.component_funcs = fu.clone();
    }
}

pub fn install_component_func_links(
    chunks: &mut [Chunk],
    links: Vec<vybe_runtime::canon_def::ComponentFuncLink>,
) {
    if links.is_empty() {
        return;
    }
    let links = Arc::new(links);
    for chunk in chunks {
        chunk.component_func_links = links.clone();
    }
}

pub fn lower_component_instance_links(
    source: Option<&vybe_ast::canon::ComponentInstanceLinks>,
) -> Result<Option<Arc<vybe_runtime::component_instance_link::InstanceLinkMetadata>>, String> {
    use vybe_runtime::component_instance_link::{InstanceExport, InstanceExports,
        InstanceFunctionImport, InstanceLinkMetadata};
    let Some(source) = source.filter(|source| !source.imports.is_empty()) else {
        return Ok(None);
    };
    let instances = source.instances.iter().map(|exports| {
        InstanceExports::new(exports.iter().map(|export| InstanceExport {
            name: Arc::from(export.item.name.as_str()),
            kind: export.item.kind,
            index: export.item.index,
            function_type: export.function_type,
        }).collect()).map_err(|error| error.message)
    }).collect::<Result<Vec<_>, _>>()?;
    let imports = source.imports.iter().map(|import| InstanceFunctionImport {
        instance: import.instance,
        name: Arc::from(import.name.as_str()),
        import: import.import,
        expected_type: import.expected_type,
    }).collect();
    Ok(Some(Arc::new(InstanceLinkMetadata { instances, imports })))
}

pub fn install_component_instance_links(
    chunks: &mut [Chunk],
    links: Option<Arc<vybe_runtime::component_instance_link::InstanceLinkMetadata>>,
) -> Result<(), String> {
    match (chunks.first_mut(), links) {
        (Some(script), links) => { script.component_instance_links = links; Ok(()) }
        (None, None) => Ok(()),
        (None, Some(_)) => Err("component instance links have no module script".into()),
    }
}

/// Fill in every `CanonCallee::CoreExport` now that the chunks exist.
///
/// `canon lift`'s `$callee` is a `core:funcidx`, and `call_canon_callee` reads
/// the number it finds there as a CHUNK index. A chunk index is assigned here,
/// during emission — the front end cannot know one, so it hands over the
/// `(class, method)` pair it *does* know and this resolves it.
///
/// Lookups use each type-table entry's method vtable. Large, repeatedly used
/// tables get borrowed-name indexes; a method is never resolved by bare name,
/// since two classes may each define the same method.
pub fn resolve_core_export_callees(
    chunks: &mut [Chunk],
    decls: &[CoreExportCallee],
    defs: &mut [CanonDef],
) -> Result<(), String> {
    // The type table lives on chunk 0.
    let types = chunks.first().map_or(&[][..], |c| c.types.as_slice());
    let core_export_count = decls.len().min(32);
    if core_export_count == 0 {
        return Ok(());
    }
    let large_type_table = types.len() >= 64;
    let large_method_table =
        core_export_count >= 32 && types.iter().any(|entry| entry.methods.len() >= 128);
    let many_core_exports = large_type_table && core_export_count >= 16;
    let type_index = if many_core_exports {
        let mut index = HashMap::with_capacity(types.len());
        for (position, entry) in types.iter().enumerate() {
            index.entry(entry.name.as_str()).or_insert(position);
        }
        Some(index)
    } else {
        None
    };
    let mut method_references = if core_export_count >= 32 && large_method_table {
        vec![0u8; types.len()]
    } else {
        Vec::new()
    };
    let mut method_index: Vec<Option<HashMap<&str, usize>>> = Vec::new();
    for d in decls {
        let class = &d.class;
        let method = &d.method;
        let position = match &type_index {
            Some(index) => index.get(class.as_str()).copied(),
            None => types.iter().position(|entry| &entry.name == class),
        }
        .ok_or_else(|| {
            let mut have: Vec<&str> = types.iter().map(|t| t.name.as_str()).collect();
            have.sort_unstable();
            format!(
                "canon {}: $callee names core instance class `{class}`, which is not in \
                 the type table (have: {})",
                d.builtin,
                if have.is_empty() {
                    "none".to_string()
                } else {
                    have.join(", ")
                }
            )
        })?;
        let entry = &types[position];
        if let Some(count) = method_references.get_mut(position) {
            if entry.methods.len() >= 128 && *count < 16 {
                *count += 1;
                if *count == 16 {
                    if method_index.is_empty() {
                        method_index.resize_with(types.len(), || None);
                    }
                    let mut methods = HashMap::with_capacity(entry.methods.len());
                    for (name, position) in &entry.methods {
                        methods.entry(name.as_str()).or_insert(*position);
                    }
                    method_index[position] = Some(methods);
                }
            }
        }
        let ci = if let Some(Some(methods)) = method_index.get(position) {
            methods.get(method.as_str()).copied()
        } else {
            entry
                .methods
                .iter()
                .find(|(name, _)| name == method)
                .map(|(_, position)| *position)
        }
        .ok_or_else(|| {
            let mut have: Vec<&str> = entry.methods.iter().map(|(m, _)| m.as_str()).collect();
            have.sort_unstable();
            format!(
                "canon {}: class `{class}` has no method `{method}` (have: {})",
                d.builtin,
                if have.is_empty() {
                    "none".to_string()
                } else {
                    have.join(", ")
                }
            )
        })?;
        let def = defs
            .get_mut(d.row)
            .ok_or_else(|| format!("canon {}: canonidx {} has no lowered row", d.builtin, d.row))?;
        match d.target {
            CoreExportTarget::Callee => {
                def.callee = Some(vybe_runtime::canon_def::CalleeRef::Core(ci as u32));
            }
            CoreExportTarget::ResourceDtor => {
                def.resource_dtor = Some(vybe_runtime::canon_def::CalleeRef::Core(ci as u32));
            }
            CoreExportTarget::Realloc => def.opts.realloc = Some(ci as u32),
            CoreExportTarget::PostReturn => def.opts.post_return = Some(ci as u32),
            CoreExportTarget::Callback => def.opts.callback = Some(ci as u32),
        }
    }
    Ok(())
}

/// The NAME a `(own <typeidx>)` / `(borrow <typeidx>)` refers to.
///
/// ⛔ Refuses rather than inventing one. `ValType::Own` binds to whatever is
/// registered under the name it carries, so a wrong or invented name binds the
/// handle to a DIFFERENT resource — which links, runs, and hands the callee
/// someone else's object. An index that names a non-resource is a different
/// mistake from one that is out of range, so they get different messages.
fn resource_name(i: u32, types: &[TypeDecl], what: &str) -> Result<String, String> {
    match types.get(i as usize) {
        Some(TypeDecl::Resource { name: Some(name), .. }) => Ok(name.clone()),
        Some(TypeDecl::Resource { name: None, .. }) => Err(format!(
            "`({what} {i})` names an UNNAMED resource type; a handle binds by name, so \
             the `(type $id (resource …))` it points at has to have a binder"
        )),
        Some(_) => Err(format!(
            "`({what} {i})` — typeidx {i} is declared but is not a resource type"
        )),
        None => Err(format!(
            "`({what} {i})` — typeidx {i} is not in the component type space (have {})",
            types.len()
        )),
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use vybe_ast::canon::CanonCoreFuncRef;
    use vybe_runtime::canon_def::CalleeRef;
    use vybe_runtime::chunk::TypeEntry;

    #[test]
    fn emitter_stream_types_relocate_without_rewriting_source_canonical_imports() {
        let element = ValType::Record(vec![("name".into(), ValType::String)]);
        let mut chunks = vec![Chunk::new("first"), Chunk::new("second")];
        let source = chunks[0].add_import("canon", "stream.read@0");
        let read = add_typed_stream_import(&mut chunks[0], "stream.read", element.clone(),
            CanonOpts { memory: Some(0), ..Default::default() });
        let drop = add_typed_stream_import(&mut chunks[0], "stream.drop-readable",
            element.clone(), Default::default());
        let second = add_typed_stream_import(&mut chunks[1], "stream.read", element.clone(),
            CanonOpts { memory: Some(0), ..Default::default() });
        assert_ne!(source, read);
        let original = CanonDef { opts: CanonOpts { is_async: true, ..Default::default() },
            ..Default::default() };
        let mut definitions = vec![original.clone()];
        let mut functions = vec![None];
        let mut values = vec![Some(ValType::U32)];
        resolve_emitter_imports(&mut chunks, &mut definitions, &mut functions, &mut values).unwrap();
        assert_eq!(definitions[0], original);
        assert_eq!(definitions.len(), 3);
        assert_eq!(values, [Some(ValType::U32), Some(ValType::Stream(Box::new(element)))]);
        assert_eq!(functions.len(), values.len());
        assert_eq!(chunks[0].imports[source as usize].name, "stream.read@0");
        assert_eq!(chunks[0].imports[read as usize].name, "stream.read@1");
        assert_eq!(chunks[0].imports[drop as usize].name, "stream.drop-readable@2");
        assert_eq!(chunks[1].imports[second as usize].name, "stream.read@1");
        assert_eq!(definitions[1].ty, Some(1));
        assert_eq!(definitions[2].ty, Some(1));
        assert!(definitions[2].opts.memory.is_none());
        assert_eq!(chunks[0].add_import("canon", "stream.read@1"), read);
        resolve_emitter_imports(&mut chunks, &mut definitions, &mut functions, &mut values).unwrap();
        assert_eq!(definitions.len(), 3);
        install_canon_section(&mut chunks, definitions);
        install_type_space(&mut chunks, functions, values, Vec::new());
        assert!(Arc::ptr_eq(&chunks[0].canon_valtypes, &chunks[1].canon_valtypes));
        assert!(chunks.iter().flat_map(|chunk| &chunk.imports)
            .all(|import| !import.name.contains(EMITTER_ROW)));
    }

    #[test]
    fn emitter_stream_type_resolution_rejects_missing_metadata() {
        let mut chunk = Chunk::new("invalid");
        chunk.add_import("canon", "stream.read@emitter:0");
        assert!(resolve_emitter_imports(&mut [chunk], &mut Vec::new(),
            &mut Vec::new(), &mut Vec::new()).unwrap_err().contains("missing emitter canonical row"));
    }

    #[test]
    fn emitter_stream_types_preserve_function_and_resource_type_slots() {
        let mut chunk = Chunk::new("typed");
        let import = add_typed_stream_import(&mut chunk, "stream.drop-readable",
            ValType::String, Default::default());
        let function = CanonFuncType { result: Some(ValType::U32), ..Default::default() };
        let mut functions = vec![Some(function.clone()), None];
        let mut values = vec![None, None];
        let mut definitions = Vec::new();
        resolve_emitter_imports(std::slice::from_mut(&mut chunk), &mut definitions,
            &mut functions, &mut values).unwrap();
        assert_eq!(functions, [Some(function), None, None]);
        assert_eq!(values, [None, None, Some(ValType::Stream(Box::new(ValType::String)))]);
        assert_eq!(definitions[0].ty, Some(2));
        assert_eq!(chunk.imports[import as usize].name, "stream.drop-readable@0");
    }

    #[test]
    fn instance_link_handoff_preserves_typed_exports_and_root_ownership() {
        use vybe_ast::canon::{ComponentInstanceExport, ComponentInstanceFunctionImport,
            ComponentInstanceLinks, ComponentItemKind, ComponentItemRef};
        assert!(lower_component_instance_links(None).unwrap().is_none());
        assert!(lower_component_instance_links(Some(&ComponentInstanceLinks::default()))
            .unwrap().is_none());
        let source = ComponentInstanceLinks {
            instances: vec![vec![
                ComponentInstanceExport { item: ComponentItemRef {
                    name: "value".into(), kind: ComponentItemKind::Func, index: 3,
                }, function_type: Some(2) },
                ComponentInstanceExport { item: ComponentItemRef {
                    name: "data".into(), kind: ComponentItemKind::Type, index: 5,
                }, function_type: None },
            ]],
            imports: vec![ComponentInstanceFunctionImport {
                instance: 0, name: "value".into(), import: 4, expected_type: 2,
            }],
        };
        let links = lower_component_instance_links(Some(&source)).unwrap().unwrap();
        assert_eq!(links.instances[0].exports()[1].kind, ComponentItemKind::Type);
        assert_eq!(links.imports[0].import, 4);
        let mut chunks = vec![Chunk::new("<script>"), Chunk::new("callee")];
        install_component_instance_links(&mut chunks, Some(links.clone())).unwrap();
        assert!(Arc::ptr_eq(chunks[0].component_instance_links.as_ref().unwrap(), &links));
        assert!(chunks[1].component_instance_links.is_none());
        assert!(install_component_instance_links(&mut [], Some(links)).is_err());
    }

    #[test]
    fn worker_parameter_does_not_resolve_in_the_component_type_space() {
        for (param, expected) in [
            (vybe_ast::canon::ComponentCoreValType::I32, vybe_runtime::canon_flat::CoreType::I32),
            (vybe_ast::canon::ComponentCoreValType::I64, vybe_runtime::canon_flat::CoreType::I64),
        ] {
            let mut row = CanonDecl::new("thread.new-indirect", None);
            row.ty = Some(99);
            row.table = Some(0);
            row.worker_param = Some(param);
            assert_eq!(lower_decl(&row).unwrap().worker_param, Some(expected));
            row.worker_param = None;
            assert!(lower_decl(&row).is_err());
        }
    }

    #[test]
    fn resource_drop_resolves_declared_core_destructor() {
        let mut drop = CanonDecl::new("resource.drop", None);
        drop.ty = Some(0);
        let decls = [drop];
        let types = [TypeDecl::Resource {
            name: Some("r".into()),
            rep: ResourceRep::I32,
            dtor: Some(CanonCallee::CoreExport {
                class: "ResourceModule".into(),
                method: "destroy".into(),
            }),
        }];
        let mut defs = lower_section(&decls).unwrap();
        let pending = attach_resource_metadata(&decls, &types, &mut defs).unwrap();
        assert_eq!(pending.len(), 1);
        assert!(defs[0].resource_dtor.is_none());
        let mut script = Chunk::new("script");
        script.types.push(TypeEntry {
            name: "ResourceModule".into(),
            methods: vec![("destroy".into(), 2)],
            ..Default::default()
        });
        resolve_core_export_callees(&mut [script], &pending, &mut defs).unwrap();
        assert_eq!(defs[0].resource_dtor, Some(CalleeRef::Core(2)));
    }

    #[test]
    fn resource_representation_width_is_attached_to_all_resource_builtins() {
        let decls: Vec<_> = ["resource.new", "resource.rep", "resource.drop"]
            .into_iter()
            .map(|name| {
                let mut decl = CanonDecl::new(name, None);
                decl.ty = Some(0);
                decl
            })
            .collect();
        let types = [TypeDecl::Resource {
            name: Some("wide".into()),
            rep: ResourceRep::I64,
            dtor: None,
        }];
        let mut defs = lower_section(&decls).unwrap();
        assert!(attach_resource_metadata(&decls, &types, &mut defs)
            .unwrap()
            .is_empty());
        assert!(defs.iter().all(|def| {
            def.resource_rep == Some(vybe_runtime::canon_flat::CoreType::I64)
        }));
    }

    #[test]
    fn context_slot_immediates_follow_supported_component_profile() {
        let mut decl = CanonDecl::new("context.get", None);
        let error = lower_decl(&decl).unwrap_err();
        assert!(error.contains("missing core value type and slot immediate"), "{error}");
        decl.context = Some(("i32".into(), 1));
        assert_eq!(lower_decl(&decl).unwrap().context, Some((ValType::I32, 1)));

        decl.context = Some(("i64".into(), 0));
        assert_eq!(lower_decl(&decl).unwrap().context, Some((ValType::I64, 0)));

        for spelling in ["f32", "f64", "u32", "bool", "string"] {
            decl.context = Some((spelling.into(), 0));
            let error = lower_decl(&decl).unwrap_err();
            assert!(
                error.contains("i32 or i64 core value type"),
                "{error}"
            );
        }

        decl.context = Some(("i32".into(), 2));
        let error = lower_decl(&decl).unwrap_err();
        assert!(error.contains("slot 2 must be less than 2"), "{error}");

        let mut other = CanonDecl::new("context.set", None);
        decl.context = Some(("i32".into(), 0));
        other.context = Some(("i64".into(), 1));
        let error = lower_section(&[decl.clone(), other.clone()]).unwrap_err();
        assert!(
            error.contains("must use the same core value type"),
            "{error}"
        );
        other.context = Some(("i32".into(), 1));
        assert_eq!(lower_section(&[decl, other]).unwrap().len(), 2);

        let mut unrelated = CanonDecl::new("thread.index", None);
        unrelated.context = Some(("i32".into(), 0));
        let error = lower_decl(&unrelated).unwrap_err();
        assert!(error.contains("unexpected context type and slot immediate"), "{error}");
    }

    #[test]
    fn installing_canon_metadata_moves_vectors_and_shares_them_across_chunks() {
        let mut chunks = vec![Chunk::new("script"), Chunk::new("body")];
        let section = vec![CanonDef::default()];
        let section_ptr = section.as_ptr();
        install_canon_section(&mut chunks, section);
        assert_eq!(chunks[0].canon_section.as_ptr(), section_ptr);
        assert!(Arc::ptr_eq(
            &chunks[0].canon_section,
            &chunks[1].canon_section
        ));

        let funcs = vec![Some(7)];
        let functypes = vec![Some(CanonFuncType::default())];
        let valtypes = vec![Some(ValType::String)];
        let funcs_ptr = funcs.as_ptr();
        let functypes_ptr = functypes.as_ptr();
        let valtypes_ptr = valtypes.as_ptr();
        install_type_space(&mut chunks, functypes, valtypes, funcs);
        assert_eq!(chunks[0].component_funcs.as_ptr(), funcs_ptr);
        assert_eq!(chunks[0].canon_functypes.as_ptr(), functypes_ptr);
        assert_eq!(chunks[0].canon_valtypes.as_ptr(), valtypes_ptr);
        assert!(Arc::ptr_eq(
            &chunks[0].component_funcs,
            &chunks[1].component_funcs
        ));
        assert!(Arc::ptr_eq(
            &chunks[0].canon_functypes,
            &chunks[1].canon_functypes
        ));
        assert!(Arc::ptr_eq(
            &chunks[0].canon_valtypes,
            &chunks[1].canon_valtypes
        ));
    }

    #[test]
    fn value_type_lowering_rejects_recursive_references() {
        let self_reference = [TypeDecl::Value(ValSpec::Ref(0))];
        let error = lower_types(&self_reference).unwrap_err();
        assert!(error.contains("recursive value type"), "{error}");

        let mutual_reference = [
            TypeDecl::Value(ValSpec::Ref(1)),
            TypeDecl::Value(ValSpec::Ref(0)),
        ];
        let error = lower_types(&mutual_reference).unwrap_err();
        assert!(error.contains("recursive value type"), "{error}");
        let error = lower_valspec(&ValSpec::Ref(0), &mutual_reference).unwrap_err();
        assert!(error.contains("recursive value type"), "{error}");
    }

    #[test]
    fn value_type_lowering_preserves_shared_forward_references() {
        let types = [
            TypeDecl::Func {
                params: vec![
                    ("left".into(), ValSpec::Ref(2)),
                    ("right".into(), ValSpec::Ref(2)),
                ],
                result: Some(ValSpec::Ref(2)),
                is_async: false,
            },
            TypeDecl::Opaque("componenttype".into()),
            TypeDecl::Value(ValSpec::Record(vec![(
                "count".into(),
                ValSpec::Prim("u32".into()),
            )])),
        ];
        let (funcs, values) = lower_types(&types).unwrap();
        let expected = ValType::Record(vec![("count".into(), ValType::U32)]);
        let function = funcs[0].as_ref().unwrap();
        assert_eq!(function.params, vec![expected.clone(), expected.clone()]);
        assert_eq!(function.result, Some(expected.clone()));
        assert!(!function.is_async);
        assert_eq!(values[2], Some(expected));
    }

    #[test]
    fn async_function_type_survives_lowering() {
        let types = [TypeDecl::Func {
            params: vec![("value".into(), ValSpec::Prim("u32".into()))],
            result: Some(ValSpec::Prim("u32".into())),
            is_async: true,
        }];
        let (functions, _) = lower_types(&types).unwrap();
        let function = functions[0].as_ref().unwrap();
        assert!(function.is_async);
        assert_eq!(function.params, vec![ValType::U32]);
        assert_eq!(function.result, Some(ValType::U32));
    }

    #[test]
    fn unsigned_type_lowering_keeps_component_signedness() {
        let types = [
            TypeDecl::Value(ValSpec::Prim("s32".into())),
            TypeDecl::Value(ValSpec::Prim("u32".into())),
            TypeDecl::Value(ValSpec::Prim("s64".into())),
            TypeDecl::Value(ValSpec::Prim("u64".into())),
        ];
        let (_, values) = lower_types(&types).unwrap();
        assert_eq!(values, vec![
            Some(ValType::I32), Some(ValType::U32),
            Some(ValType::I64), Some(ValType::U64),
        ]);
    }

    #[test]
    fn value_type_lowering_preserves_invalid_reference_errors() {
        let types = [
            TypeDecl::Resource { name: Some("r".into()), rep: ResourceRep::I32, dtor: None },
            TypeDecl::Func {
                params: Vec::new(),
                result: None,
                is_async: false,
            },
            TypeDecl::Opaque("componenttype".into()),
        ];
        for (index, expected) in [
            (0, "RESOURCE used directly"),
            (1, "FUNCTION type used as a value type"),
            (2, "componenttype"),
            (3, "not in the component type space"),
        ] {
            let error = lower_valspec(&ValSpec::Ref(index), &types).unwrap_err();
            assert!(error.contains(expected), "{error}");
        }
    }

    #[test]
    fn core_export_resolution_borrows_type_table_and_preserves_errors() {
        let mut chunk = Chunk::new("script");
        chunk.types.push(TypeEntry {
            name: "Counter".into(),
            methods: vec![("get".into(), 7)],
            ..Default::default()
        });
        let mut decl = CanonDecl::new("lift", None);
        decl.callee = Some(CanonCallee::CoreExport {
            class: "Counter".into(),
            method: "get".into(),
        });
        let mut defs = lower_section(std::slice::from_ref(&decl)).unwrap();
        let mut chunks = vec![chunk];

        resolve_core_export_callees(
            &mut chunks,
            &collect_core_export_callees(std::slice::from_ref(&decl)),
            &mut defs,
        )
        .unwrap();
        assert_eq!(defs[0].callee, Some(CalleeRef::Core(7)));
        assert_eq!(chunks[0].types[0].methods[0].1, 7);

        decl.callee = Some(CanonCallee::CoreExport {
            class: "Counter".into(),
            method: "missing".into(),
        });
        let error = resolve_core_export_callees(
            &mut chunks,
            &collect_core_export_callees(&[decl]),
            &mut defs,
        )
        .unwrap_err();
        assert!(error.contains("has no method `missing`"), "{error}");
    }

    #[test]
    fn canonical_options_resolve_exports_without_changing_numeric_indices() {
        let mut chunk = Chunk::new("script");
        chunk.types.push(TypeEntry {
            name: "MemoryModule".into(),
            methods: vec![
                ("allocate".into(), 11),
                ("finish".into(), 13),
                ("resume".into(), 17),
            ],
            ..Default::default()
        });
        let mut decl = CanonDecl::new("lift", None);
        let export = |method: &str| CanonCoreFuncRef::CoreExport {
            class: "MemoryModule".into(),
            method: method.into(),
        };
        decl.opts.realloc = Some(export("allocate"));
        decl.opts.post_return = Some(export("finish"));
        decl.opts.callback = Some(export("resume"));
        let mut defs = lower_section(std::slice::from_ref(&decl)).unwrap();
        assert_eq!(defs[0].opts.realloc, None);
        let mut chunks = vec![chunk];
        resolve_core_export_callees(
            &mut chunks,
            &collect_core_export_callees(std::slice::from_ref(&decl)),
            &mut defs,
        ).unwrap();
        assert_eq!(defs[0].opts.realloc, Some(11));
        assert_eq!(defs[0].opts.post_return, Some(13));
        assert_eq!(defs[0].opts.callback, Some(17));

        decl.opts.callback = Some(export("missing"));
        let mut missing_defs = lower_section(std::slice::from_ref(&decl)).unwrap();
        let error = resolve_core_export_callees(
            &mut chunks,
            &collect_core_export_callees(std::slice::from_ref(&decl)),
            &mut missing_defs,
        ).unwrap_err();
        assert!(error.contains("has no method `missing`"), "{error}");

        decl.opts.realloc = Some(CanonCoreFuncRef::Index(5));
        decl.opts.post_return = None;
        decl.opts.callback = None;
        let defs = lower_section(&[decl]).unwrap();
        assert_eq!(defs[0].opts.realloc, Some(5));
    }

    #[test]
    fn sparse_core_export_callees_keep_original_canon_indices() {
        let mut decls = vec![
            CanonDecl::new("stream.read", None),
            CanonDecl::new("lower", None),
            CanonDecl::new("lift", None),
        ];
        decls[2].callee = Some(CanonCallee::CoreExport {
            class: "Counter".into(),
            method: "get".into(),
        });
        let pending = collect_core_export_callees(&decls);
        assert_eq!(pending.len(), 1);
        assert_eq!(pending[0].row, 2);
        let mut defs = lower_section(&decls).unwrap();
        let mut chunk = Chunk::new("script");
        chunk.types.push(TypeEntry {
            name: "Counter".into(),
            methods: vec![("get".into(), 5)],
            ..Default::default()
        });
        let mut chunks = vec![chunk];
        resolve_core_export_callees(&mut chunks, &pending, &mut defs).unwrap();
        assert_eq!(defs[0].callee, None);
        assert_eq!(defs[1].callee, None);
        assert_eq!(defs[2].callee, Some(CalleeRef::Core(5)));

        defs.pop();
        let error = resolve_core_export_callees(&mut chunks, &pending, &mut defs).unwrap_err();
        assert!(error.contains("canon lift: canonidx 2"), "{error}");
    }

    #[test]
    fn core_export_resolution_indexes_large_tables_without_changing_first_match() {
        let mut chunk = Chunk::new("script");
        for position in 0..80 {
            chunk.types.push(TypeEntry {
                name: format!("Type{position}"),
                methods: vec![("run".into(), position)],
                ..Default::default()
            });
        }
        chunk.types.push(TypeEntry {
            name: "Type79".into(),
            methods: vec![("run".into(), 999)],
            ..Default::default()
        });
        let positions = [79, 4, 28, 0, 71, 3, 55, 9, 64, 15, 38, 22, 49, 6, 73, 1];
        let mut decls: Vec<_> = positions
            .iter()
            .copied()
            .map(|position| {
                let mut decl = CanonDecl::new("lift", None);
                decl.callee = Some(CanonCallee::CoreExport {
                    class: format!("Type{position}"),
                    method: "run".into(),
                });
                decl
            })
            .collect();
        let mut defs = lower_section(&decls).unwrap();
        let mut chunks = vec![chunk];

        resolve_core_export_callees(&mut chunks, &collect_core_export_callees(&decls), &mut defs)
            .unwrap();
        let callees: Vec<_> = defs.iter().map(|d| d.callee).collect();
        let expected: Vec<_> = positions
            .iter()
            .map(|&i| Some(CalleeRef::Core(i as u32)))
            .collect();
        assert_eq!(callees, expected);

        decls[0].callee = Some(CanonCallee::CoreExport {
            class: "Missing".into(),
            method: "run".into(),
        });
        let error = resolve_core_export_callees(
            &mut chunks,
            &collect_core_export_callees(&decls),
            &mut defs,
        )
        .unwrap_err();
        assert!(error.contains("class `Missing`"), "{error}");
    }

    #[test]
    fn core_export_resolution_indexes_repeated_methods_without_changing_first_match() {
        let mut chunk = Chunk::new("script");
        let mut methods: Vec<_> = (0..256)
            .map(|position| (format!("method{position}"), position))
            .collect();
        methods.push(("method255".into(), 999));
        chunk.types.push(TypeEntry {
            name: "Large".into(),
            methods,
            ..Default::default()
        });
        let mut decls: Vec<_> = (0..64)
            .map(|position| {
                let mut decl = CanonDecl::new("lift", None);
                decl.callee = Some(CanonCallee::CoreExport {
                    class: "Large".into(),
                    method: format!("method{}", position * 4 + 3),
                });
                decl
            })
            .collect();
        decls[0].callee = Some(CanonCallee::CoreExport {
            class: "Large".into(),
            method: "method255".into(),
        });
        let mut defs = lower_section(&decls).unwrap();
        let mut chunks = vec![chunk];
        resolve_core_export_callees(&mut chunks, &collect_core_export_callees(&decls), &mut defs)
            .unwrap();
        assert_eq!(defs[0].callee, Some(CalleeRef::Core(255)));
        assert_eq!(defs[1].callee, Some(CalleeRef::Core(7)));

        decls[0].callee = Some(CanonCallee::CoreExport {
            class: "Large".into(),
            method: "missing".into(),
        });
        let error = resolve_core_export_callees(
            &mut chunks,
            &collect_core_export_callees(&decls),
            &mut defs,
        )
        .unwrap_err();
        assert!(error.contains("has no method `missing`"), "{error}");
    }

    #[test]
    fn core_export_method_index_keeps_first_class_with_duplicate_name() {
        let mut chunk = Chunk::new("script");
        chunk.types.push(TypeEntry {
            name: "Same".into(),
            methods: vec![("run".into(), 11)],
            ..Default::default()
        });
        chunk.types.push(TypeEntry {
            name: "Same".into(),
            methods: (0..256)
                .map(|position| (format!("method{position}"), position))
                .chain(std::iter::once(("run".into(), 99)))
                .collect(),
            ..Default::default()
        });
        let decls: Vec<_> = (0..32)
            .map(|_| {
                let mut decl = CanonDecl::new("lift", None);
                decl.callee = Some(CanonCallee::CoreExport {
                    class: "Same".into(),
                    method: "run".into(),
                });
                decl
            })
            .collect();
        let mut defs = lower_section(&decls).unwrap();
        resolve_core_export_callees(
            &mut [chunk],
            &collect_core_export_callees(&decls),
            &mut defs,
        )
        .unwrap();
        assert!(defs
            .iter()
            .all(|def| def.callee == Some(CalleeRef::Core(11))));
    }
}

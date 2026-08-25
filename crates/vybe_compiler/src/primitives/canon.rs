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

use std::sync::Arc;
use vybe_ast::canon::{CanonCallee, CanonDecl, TypeDecl, ValSpec};
use vybe_runtime::canon_def::{CanonDef, CanonFuncType, CanonOpts, CanonSection, StringEncoding};
use vybe_runtime::chunk::Chunk;
use vybe_runtime::component::ValType;

/// The primitive a `context.get` / `context.set` slot holds.
///
/// Refuses a spelling it cannot name rather than substituting one. The width of
/// a context slot decides how many bytes move; a guess corrupts the task's
/// context with nothing to detect it.
fn context_valtype(spelling: &str) -> Result<ValType, String> {
    match spelling {
        "i32" | "s32" | "u32" => Ok(ValType::I32),
        "i64" | "s64" | "u64" => Ok(ValType::I64),
        "f64" | "f32" => Ok(ValType::F64),
        "bool" => Ok(ValType::Bool),
        "string" => Ok(ValType::String),
        other => Err(format!(
            "canon context.*: `{other}` is not a primitive this context slot can hold"
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
        ty: d.ty,
        functype: d.functype,
        results: d.results.clone(),
        context: match &d.context {
            Some((v, i)) => Some((context_valtype(v)?, *i)),
            None => None,
        },
        table: d.table,
        opts: CanonOpts {
            string_encoding: string_encoding(&d.opts.string_encoding)?,
            memory: d.opts.memory,
            realloc: d.opts.realloc,
            post_return: d.opts.post_return,
            callback: d.opts.callback,
            is_async: d.opts.is_async,
        },
        cancellable: d.cancellable,
        is_async: d.is_async,
        shared: d.shared,
    })
}

/// Lower a whole declared section, preserving canonidx order.
pub fn lower_section(decls: &[CanonDecl]) -> Result<Vec<CanonDef>, String> {
    decls
        .iter()
        .map(|d| {
            lower_decl(d).map_err(|e| {
                // Quote the row as the source spelled it — a canonidx alone
                // names nothing a reader can find.
                format!("canon {}: {e}", d.builtin)
            })
        })
        .collect()
}

/// Publish the section to every chunk, exactly as `normalize_global_table`
/// publishes the global index space.
///
/// Every chunk, not just chunk 0: a chunk that carries a canonidx has to be
/// able to say what that index means, and threading the table through every
/// caller is what dragged `cli.rs` in when the global table was parked on
/// chunk 0.
pub fn install_canon_section(chunks: &mut [Chunk], section: &[CanonDef]) {
    if section.is_empty() {
        return;
    }
    let shared: CanonSection = Arc::new(section.to_vec());
    for chunk in chunks.iter_mut() {
        chunk.canon_section = shared.clone();
    }
}

/// `vybe_ast::canon::ValSpec` → `vybe_runtime::component::ValType`.
///
/// Refuses every spelling the runtime cannot represent FAITHFULLY, rather than
/// mapping it onto the nearest thing that compiles.
///
/// The narrow widths, `f32` and `char` are now real `ValType`s and lower
/// directly. ⛔ **`s32`/`u32` and `s64`/`u64` still collapse to `I32`/`I64`,
/// and that remains a deviation**: `flatten_type` sends both signednesses to
/// one core type, so nothing is lost FLAT — but `load` spells them as separate
/// cases (`load_int(cx, ptr, 4)` versus `signed = True`), so a `u32` holding
/// `0xFFFFFFFF` should lift as `4294967295` and lifts as `-1`. Splitting them
/// is a rename across every platform that constructs `ValType::I32`, which is
/// why it is recorded in cmplan.md rather than done here.
pub fn lower_valspec(v: &ValSpec, types: &[TypeDecl]) -> Result<ValType, String> {
    Ok(match v {
        ValSpec::Prim(p) => match p.as_str() {
            "bool" => ValType::Bool,
            "s32" | "u32" => ValType::I32,
            "s64" | "u64" => ValType::I64,
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
        ValSpec::List(t) => ValType::List(Box::new(lower_valspec(t, types)?)),
        ValSpec::Option(t) => ValType::Option(Box::new(lower_valspec(t, types)?)),
        ValSpec::Result(ok, err) => ValType::Result(
            match ok {
                Some(t) => Some(Box::new(lower_valspec(t, types)?)),
                None => None,
            },
            match err {
                Some(t) => Some(Box::new(lower_valspec(t, types)?)),
                None => None,
            },
        ),
        ValSpec::Record(fields) => {
            let mut out = Vec::with_capacity(fields.len());
            for (n, t) in fields {
                out.push((n.clone(), lower_valspec(t, types)?));
            }
            ValType::Record(out)
        }
        ValSpec::Variant(cases) => {
            let mut out = Vec::with_capacity(cases.len());
            for (n, t) in cases {
                out.push((
                    n.clone(),
                    match t {
                        Some(t) => Some(lower_valspec(t, types)?),
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
        ValSpec::Stream(Some(t)) => ValType::Stream(Box::new(lower_valspec(t, types)?)),
        ValSpec::Future(Some(t)) => ValType::Future(Box::new(lower_valspec(t, types)?)),
        ValSpec::Stream(None) => {
            return Err(
                "`(stream)` with no element type — `ValType::Stream` has nowhere to \
                 record the absence, and a stand-in element would feed `elem_size`"
                    .to_string(),
            )
        }
        ValSpec::Future(None) => {
            return Err(
                "`(future)` with no element type — `ValType::Future` has nowhere to \
                 record the absence, and a stand-in element would feed `elem_size`"
                    .to_string(),
            )
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
                out.push((i.to_string(), lower_valspec(t, types)?));
            }
            ValType::Record(out)
        }
        ValSpec::Enum(labels) => {
            ValType::Variant(labels.iter().map(|l| (l.clone(), None)).collect())
        }
        ValSpec::Map(k, v) => ValType::List(Box::new(ValType::Record(vec![
            ("0".to_string(), lower_valspec(k, types)?),
            ("1".to_string(), lower_valspec(v, types)?),
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
            ValType::ListFixed(Box::new(lower_valspec(elem, types)?), *n)
        }
        // ⛔ The runtime names a resource by STRING and the source names it by
        // INDEX. `TypeDecl::Resource(name)` is the only record connecting the
        // two, and it is filled from the `(type $r …)` binder.
        ValSpec::Own(i) => ValType::Own(resource_name(*i, types, "own")?),
        ValSpec::Borrow(i) => ValType::Borrow(resource_name(*i, types, "borrow")?),
        // A typeidx used AS a value type — `(type $t u32)` then `(param "a" $t)`.
        // It resolves to whatever that index declares, and refuses if the index
        // declares something that is not a value.
        ValSpec::Ref(i) => match types.get(*i as usize) {
            Some(TypeDecl::Value(inner)) => lower_valspec(inner, types)?,
            Some(TypeDecl::Resource(_)) => {
                return Err(format!(
                    "`(type {i})` names a RESOURCE used directly as a value type; a \
                     resource travels as `(own {i})` or `(borrow {i})`, never bare"
                ))
            }
            Some(TypeDecl::Func { .. }) => {
                return Err(format!(
                    "`(type {i})` names a FUNCTION type used as a value type"
                ))
            }
            Some(TypeDecl::Opaque(k)) => {
                return Err(format!("`(type {i})` names a `{k}`, which is not a value type"))
            }
            None => {
                return Err(format!(
                    "`(type {i})` is not in the component type space (have {})",
                    types.len()
                ))
            }
        },
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
    for (i, t) in types.iter().enumerate() {
        match t {
            TypeDecl::Func { params, result } => {
                let mut ps = Vec::with_capacity(params.len());
                for (n, v) in params {
                    ps.push(lower_valspec(v, types).map_err(|e| {
                        format!("component type {i}: param \"{n}\": {e}")
                    })?);
                }
                let r = match result {
                    Some(v) => Some(
                        lower_valspec(v, types)
                            .map_err(|e| format!("component type {i}: result: {e}"))?,
                    ),
                    None => None,
                };
                functypes.push(Some(CanonFuncType {
                    params: ps,
                    result: r,
                }));
                valtypes.push(None);
            }
            TypeDecl::Value(v) => {
                functypes.push(None);
                valtypes.push(Some(
                    lower_valspec(v, types).map_err(|e| format!("component type {i}: {e}"))?,
                ));
            }
            // Occupies its index and nothing more — which is the point: it must
            // still advance the space, or every later typeidx shifts.
            // A resource occupies its index and is neither a function type
            // nor a value type — a handle to it is `own`/`borrow`.
            TypeDecl::Resource(_) => {
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
    functypes: &[Option<CanonFuncType>],
    valtypes: &[Option<ValType>],
    funcs: &[Option<u32>],
) {
    if functypes.is_empty() && valtypes.is_empty() && funcs.is_empty() {
        return;
    }
    let fu: Arc<Vec<Option<u32>>> = Arc::new(funcs.to_vec());
    let f: Arc<Vec<Option<CanonFuncType>>> = Arc::new(functypes.to_vec());
    let v: Arc<Vec<Option<ValType>>> = Arc::new(valtypes.to_vec());
    for chunk in chunks.iter_mut() {
        chunk.canon_functypes = f.clone();
        chunk.canon_valtypes = v.clone();
        chunk.component_funcs = fu.clone();
    }
}

/// Fill in every `CanonCallee::CoreExport` now that the chunks exist.
///
/// `canon lift`'s `$callee` is a `core:funcidx`, and `call_canon_callee` reads
/// the number it finds there as a CHUNK index. A chunk index is assigned here,
/// during emission — the front end cannot know one, so it hands over the
/// `(class, method)` pair it *does* know and this resolves it.
///
/// The lookup is the type table's vtable, which is already exactly this map:
/// `TypeEntry { name, methods: Vec<(String, usize)> }`. Nothing new is built,
/// and in particular nothing resolves a method by BARE NAME — two classes may
/// each have a `run`, and the debugger's chunk list shows several same-named
/// chunks for exactly that reason.
pub fn resolve_core_export_callees(
    chunks: &mut [Chunk],
    decls: &[CanonDecl],
    defs: &mut [CanonDef],
) -> Result<(), String> {
    // The type table lives on chunk 0.
    let types = match chunks.first() {
        Some(c) => c.types.clone(),
        None => Vec::new(),
    };
    for (i, d) in decls.iter().enumerate() {
        let Some(CanonCallee::CoreExport { class, method }) = &d.callee else {
            continue;
        };
        let entry = types.iter().find(|t| &t.name == class).ok_or_else(|| {
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
        let ci = entry
            .methods
            .iter()
            .find(|(m, _)| m == method)
            .map(|(_, ci)| *ci)
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
        let def = defs.get_mut(i).ok_or_else(|| {
            format!("canon {}: canonidx {i} has no lowered row", d.builtin)
        })?;
        def.callee = Some(vybe_runtime::canon_def::CalleeRef::Core(ci as u32));
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
        Some(TypeDecl::Resource(Some(name))) => Ok(name.clone()),
        Some(TypeDecl::Resource(None)) => Err(format!(
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

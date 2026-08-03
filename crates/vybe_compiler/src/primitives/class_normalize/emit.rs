//! `emit_class` — compile a `NormalClass` to bytecode + register its
//! runtime `ClassType` in the `TypeRegistry`.
//!
//! This function is intentionally language-neutral. Every source-language
//! quirk was resolved by the walker's `normalize_class` pass before
//! reaching here. Read `types.rs` and `canonical.rs` for the data
//! contract; `classnormalization.md` for the full architecture.
//!
//! # Phase status (Phase 2b.1 — consolidated shim)
//!
//! After Phase 3 landed all 8 languages on the `normalize_class` path,
//! this file collapsed to two thin functions: `emit_class_from_ast`
//! (dispatches to the per-language normalizer) and `emit_class` (hands
//! the `NormalClass` to `Compiler::compile_class`). The
//! `NormalClass → ClassMember` reconstruction that the legacy
//! orchestration still consumes now lives inside `compile_class`
//! itself — see `reconstruct_members_for_compile` in
//! `primitives/classes.rs`. Phase 2b.2 eliminates that reconstruction
//! by porting each pass of `compile_class` to read `NormalClass`
//! fields directly.

use crate::ast::{ClassMember, ClassModifiers, Span, StmtKind};
use crate::primitives::Compiler;
use crate::primitives::class_normalize::{access_from_visibility, from_method_stmt};
use vybe_ast::class_normalize::types::*;

/// Entry point from `compile_stmt`. Receives the raw AST fields from
/// `StmtKind::ClassDecl`, normalises per language, then hands off to
/// `emit_class`.
/// Produce the normalized class WITHOUT emitting it.
///
/// Split out of `emit_class_from_ast` so the declaration pass can normalize
/// every class up front (`link.rs`), instead of normalization happening once
/// per class during code generation. That ordering is what makes a class's
/// member set — and its declared augmentations — knowable before any body
/// compiles. See flexclassplan.md §3a, §4c.
pub fn normalize_class_from_ast(
    compiler: &Compiler,
    span: Span,
    cname: &str,
    parents: &[String],
    interfaces: &[String],
    members: &[ClassMember],
    modifiers: &ClassModifiers,
    is_value_type: bool,
) -> Result<NormalClass, String> {
    let mut nc = if compiler.profile.uses_normalize_class {
        normalize_for_profile(
            compiler.profile.name.as_str(),
            span.clone(),
            cname,
            parents,
            interfaces,
            members,
            modifiers,
        )?
    } else {
        normalize_from_ast_legacy(span.clone(), cname, parents, interfaces, members, modifiers)
    };
    nc.is_value_type = is_value_type;
    // Fill the fields that are a straight copy of this function's own
    // arguments. Every one of the twelve normalizers wrote these identically —
    // `parent: parents.first().cloned()`, `is_abstract: modifiers.is_abstract`,
    // … — which is duplication AND a place to silently disagree (a language
    // could take `parents.last()` and nothing would catch it). The caller
    // already holds the answers, so it states them once.
    //
    nc.span = span;
    nc.name = cname.to_string();
    nc.parent = parents.first().cloned();
    nc.is_abstract = modifiers.is_abstract;
    nc.is_sealed = modifiers.is_sealed;
    nc.is_partial = modifiers.is_partial;
    // ALL direct bases: `parent` is `bases[0]`; `bases[1..]` is only read
    // behind the `class_multiple_inheritance` opt-in.
    nc.bases = parents.to_vec();
    // Interfaces are a COMMON concept — `classes.rs` merges them with
    // reflection interfaces, dedups and registers them, which is what lets a
    // Java class implement a PHP interface. So the DECLARED list is filled here
    // and a normalizer never copies it.
    //
    // A language may still ADD an entry its rules mandate (Pascal: every class
    // implicitly roots at `TObject`; Python: extra bases join the list under
    // multiple inheritance). Those append to the declared list rather than
    // replacing it — additive rules, not a second implementation.
    let mut merged = interfaces.to_vec();
    for extra in std::mem::take(&mut nc.interfaces) {
        if !merged.iter().any(|i| i.eq_ignore_ascii_case(&extra)) {
            merged.push(extra);
        }
    }
    nc.interfaces = merged;
    Ok(nc)
}

pub fn emit_class_from_ast(
    compiler: &mut Compiler,
    span: Span,
    cname: &str,
    parents: &[String],
    interfaces: &[String],
    members: &[ClassMember],
    modifiers: &ClassModifiers,
    is_value_type: bool,
) -> Result<(), String> {
    // Reuse the class normalized during the DECLARATION pass when it is
    // available. That copy has had its augmentations folded in (traits /
    // mixins / promoted fields), and re-normalizing from the AST here would
    // silently discard them — the class would emit without its contributed
    // members. It is also the point of having one class model: normalize once,
    // use everywhere. See flexclassplan.md §3a, §4c.
    let canon = compiler.canon(cname);
    let mut nc = match compiler.normalized_classes.get(&canon) {
        Some(stored) => {
            let mut nc = stored.clone();
            // These two are supplied by the caller at definition time; the
            // declaration pass has no way to know them.
            nc.is_value_type = is_value_type;
            nc.bases = parents.to_vec();
            nc
        }
        None => normalize_class_from_ast(
            compiler,
            span,
            cname,
            parents,
            interfaces,
            members,
            modifiers,
            is_value_type,
        )? };
    // Set centrally, like `bases`: every path lands here — the stored
    // declaration-pass copy and each per-language normalizer alike — so no
    // language has to remember to carry it.
    nc.declared_kind = modifiers.kind;
    // Link classes for `NextInOrder` augmentations come first: this class's
    // parent IS the last of them, and a class cannot wire a prototype chain
    // through a parent that has not been defined yet. They are emitted here,
    // off the using class's declaration, because there is no `ClassDecl` of
    // their own to reach them by — the declaration pass synthesized them
    // (flexclassplan.md §4c-R).
    for link_name in nc.synthesized_bases.clone() {
        let link_key = compiler.canon(&link_name);
        let Some(link) = compiler.normalized_classes.get(&link_key).cloned() else {
            return Err(format!(
                "internal: synthesized base `{link_name}` of class `{cname}` is missing"
            ));
        };
        emit_class(compiler, link)?;
    }
    emit_class(compiler, nc)
}

fn normalize_from_ast_legacy(
    span: Span,
    name: &str,
    parents: &[String],
    interfaces: &[String],
    members: &[ClassMember],
    modifiers: &ClassModifiers,
) -> NormalClass {
    let mut instance_fields = Vec::new();
    let mut static_fields = Vec::new();
    let mut instance_methods = Vec::new();
    let mut static_methods = Vec::new();
    let mut raw_extra_members = Vec::new();

    for member in members {
        match member {
            ClassMember::Field {
                name,
                type_hint,
                init,
                modifiers,
                array_bounds,
                ..
            } => {
                let field = NormalField {
                    span: span.clone(),
                    name: name.clone(),
                    type_hint: type_hint.clone(),
                    init: init.clone(),
                    array_bounds: array_bounds.clone(),
                    access: access_from_visibility(modifiers.visibility),
                    readonly: modifiers.is_readonly };
                if modifiers.is_shared {
                    static_fields.push(field);
                } else {
                    instance_fields.push(field);
                }
            }
            ClassMember::Method(stmt) => {
                if let StmtKind::FunctionDecl {
                    name: method_name,
                    modifiers,
                    ..
                } = &stmt.kind
                {
                    if let Some(method) = from_method_stmt(
                        stmt.span.clone(),
                        stmt,
                        method_name,
                        access_from_visibility(modifiers.visibility),
                    ) {
                        if modifiers.is_shared {
                            static_methods.push(method);
                        } else {
                            instance_methods.push(method);
                        }
                    }
                } else {
                    raw_extra_members.push(member.clone());
                }
            }
            _ => raw_extra_members.push(member.clone()) }
    }

    NormalClass {
        span,
        name: name.to_string(),
        parent: parents.first().cloned(),
        interfaces: interfaces.to_vec(),
        is_abstract: modifiers.is_abstract,
        is_sealed: modifiers.is_sealed,
        is_partial: modifiers.is_partial,
        explicit_self_param: true,
        instance_fields,
        static_fields,
        instance_methods,
        static_methods,
        raw_extra_members,
        ..Default::default()
    }
}

/// Compile a `NormalClass` — the compiler-neutral entry point.
pub fn emit_class(compiler: &mut Compiler, class: NormalClass) -> Result<(), String> {
    compiler.compile_class(&class)
}

/// Dispatch `normalize_class` to the right language implementation based
/// on the profile name.
fn normalize_for_profile(
    lang: &str,
    span: Span,
    name: &str,
    parents: &[String],
    interfaces: &[String],
    members: &[ClassMember],
    modifiers: &ClassModifiers,
) -> Result<NormalClass, String> {
    // Dispatch through the language registry — no `crate::languages::<lang>`
    // paths, so a language can live in its own (eventually dylib) crate and
    // still be reached here.
    crate::ensure_languages_registered();
    vybe_runtime::registry::normalize_class(
        lang, span, name, parents, interfaces, members, modifiers,
    )
    .ok_or_else(|| {
        format!(
            "normalize_class not yet implemented for language {:?} — set \
             `uses_normalize_class = false` in the profile until Phase 3 \
             covers it",
            lang
        )
    })
}

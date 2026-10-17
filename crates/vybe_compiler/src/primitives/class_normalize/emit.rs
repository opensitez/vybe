//! `emit_class` — compile a `NormalClass` to bytecode + register its
//! runtime `ClassType` in the `TypeRegistry`.
//!
//! This function is intentionally language-neutral. Every source-language
//! quirk was resolved by the walker's `normalize_class` pass before
//! reaching here. Read `types.rs` and `canonical.rs` for the data
//! contract; `classnormalization.md` for the full architecture.
//!
//! Class compilation consumes `NormalClass` fields directly. This adapter
//! still constructs the owned normalized declaration from borrowed source AST
//! members, so moving normalization to an owned frontend reduction can remove
//! body and parameter copies at that boundary.

use crate::ast::{ClassMember, ClassModifiers, Span, StmtKind};
use crate::primitives::Compiler;
use crate::primitives::class_normalize::{access_from_visibility, from_method_stmt};
use vybe_ast::class_normalize::types::*;

fn validate_byte_layout(
    cname: &str,
    members: &[ClassMember],
    semantics: &vybe_ast::ValueSemantics,
) -> Result<(), String> {
    if let Some(layout) = &semantics.byte_layout {
        layout.validate().map_err(|error| format!("{cname}: {error}"))?;
        for field in &layout.fields {
            if !members.iter().any(|member| matches!(member,
                ClassMember::Field { name, .. } if name == &field.name)) {
                return Err(format!("{cname}: byte layout names undeclared field {}", field.name));
            }
        }
    }
    Ok(())
}

/// Entry point from `compile_stmt`. Receives the raw AST fields from
/// `StmtKind::ClassDecl`, normalises per language, then hands off to
/// `emit_class`.
/// Resolve each declared base NAME to the class identity it denotes.
///
/// A base is written in source the way the programmer sees it — `Inherits
/// Animal`, `extends Animal` — while the class it names may live in a
/// namespace and be registered as `animals.animal`. Case-folding alone
/// (`canon`) yields `animal`, which matches no class, and the derived class
/// then inherits NOTHING: its own methods cannot see an inherited member by
/// bare name, even though the base's own methods still can, and writes and
/// external reads keep working. That combination reads like a property bug
/// rather than a resolution one.
///
/// Done here because this is the single point where `parent` and `bases` are
/// filled for every language — resolving in `classes.rs` would fix `parent`
/// and leave the `bases[1..]` multiple-inheritance path on the raw spelling,
/// which is the same defect one arm over.
///
/// `resolve_source_namespace_type` is the resolver the `New` path and
/// `metadata.rs` already use, so the enclosing namespace and the file's
/// imports are honoured here exactly as they are there. An unresolved name
/// keeps its original spelling, so a base that is not a user-namespaced type
/// reaches the existing `canon` downstream unchanged.
fn resolve_declared_bases(compiler: &Compiler, parents: &[String]) -> Vec<String> {
    // PHP's walker has already resolved `extends` and `implements`, including
    // leading `\` global names. Resolving a global `InvalidArgumentException`
    // again inside namespace App would turn it into App.InvalidArgumentException
    // and make the class inherit from itself.
    if compiler.profile.name == "php" {
        return parents.to_vec();
    }
    parents
        .iter()
        .map(|p| {
            compiler
                .resolve_source_namespace_type(p)
                .unwrap_or_else(|| p.clone())
        })
        .collect()
}

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
    semantics: vybe_ast::ValueSemantics,
) -> Result<NormalClass, String> {
    normalize_class_from_ast_with_override(compiler, span, cname, parents, interfaces, members, modifiers, semantics, None)
}

pub(crate) fn normalize_class_from_ast_with_override(
    compiler: &Compiler, span: Span, cname: &str, parents: &[String], interfaces: &[String],
    members: &[ClassMember], modifiers: &ClassModifiers, semantics: vybe_ast::ValueSemantics,
    supplied: Option<NormalClass>,
) -> Result<NormalClass, String> {
    validate_byte_layout(cname, members, &semantics)?;
    let mut nc = if let Some(supplied) = supplied { supplied }
    else if compiler.profile.uses_normalize_class {
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
    // `is_value_type` is STORAGE, its original meaning: does assignment copy.
    // Equality is a separate axis and reads `nc.semantics.equality`.
    nc.is_value_type = semantics.storage == vybe_ast::ValueStorage::Value;
    nc.semantics = semantics;
    // Fill the fields that are a straight copy of this function's own
    // arguments. Every one of the twelve normalizers wrote these identically —
    // `parent: parents.first().cloned()`, `is_abstract: modifiers.is_abstract`,
    // … — which is duplication AND a place to silently disagree (a language
    // could take `parents.last()` and nothing would catch it). The caller
    // already holds the answers, so it states them once.
    //
    nc.declared_kind = modifiers.kind;
    nc.library_identity = match &modifiers.linkage {
        vybe_ast::ClassLinkage::Definition => None,
        vybe_ast::ClassLinkage::Library { identity }
        | vybe_ast::ClassLinkage::LibraryReference { identity } => Some(identity.clone()),
    };
    nc.span = span;
    nc.name = cname.to_string();
    let resolved_parents = resolve_declared_bases(compiler, parents);
    nc.parent = resolved_parents.first().cloned();
    nc.is_abstract = modifiers.is_abstract;
    nc.is_sealed = modifiers.is_sealed;
    nc.is_partial = modifiers.is_partial;
    // ALL direct bases: `parent` is `bases[0]`; `bases[1..]` is only read
    // behind the `class_multiple_inheritance` opt-in.
    nc.bases = resolved_parents;
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
    semantics: vybe_ast::ValueSemantics,
) -> Result<(), String> {
    // Its full declaration already participated in the common declaration
    // passes. The provider's existing runtime definition supplies the code.
    if matches!(modifiers.linkage, vybe_ast::ClassLinkage::LibraryReference { .. }) {
        return Ok(());
    }
    validate_byte_layout(cname, members, &semantics)?;
    // Reuse the class normalized during the DECLARATION pass when it is
    // available. That copy has had its augmentations folded in (traits /
    // mixins / promoted fields), and re-normalizing from the AST here would
    // silently discard them — the class would emit without its contributed
    // members. It is also the point of having one class model: normalize once,
    // use everywhere. See flexclassplan.md §3a, §4c.
    let canon = compiler.canon(cname);
    // Normalization and augmentation finish before emission. Keep a shared
    // registry view so method bodies can be compiled without copying them.
    let declarations = compiler.normalized_classes.clone();
    let resolved_bases = resolve_declared_bases(compiler, parents);
    let mut nc = match declarations.get(&canon) {
        Some(stored) => {
            // Variant declarations may contain expressions and have no value
            // equality contract. Keep their existing owned update path.
            let unchanged = stored.is_value_type == (semantics.storage == vybe_ast::ValueStorage::Value)
                && stored.semantics.storage == semantics.storage
                && stored.semantics.equality == semantics.equality
                && stored.semantics.layout == semantics.layout
                && stored.semantics.byte_layout == semantics.byte_layout
                && stored.semantics.variant.is_none() && semantics.variant.is_none()
                && stored.bases == resolved_bases
                && stored.declared_kind == modifiers.kind;
            if unchanged {
                std::borrow::Cow::Borrowed(stored)
            } else {
                let mut owned = stored.clone();
                owned.is_value_type = semantics.storage == vybe_ast::ValueStorage::Value;
                owned.semantics = semantics;
                owned.bases = resolved_bases;
                owned.declared_kind = modifiers.kind;
                std::borrow::Cow::Owned(owned)
            }
        }
        None => std::borrow::Cow::Owned(normalize_class_from_ast(
            compiler, span, cname, parents, interfaces, members, modifiers, semantics,
        )?),
    };
    #[cfg(test)]
    if super::super::COPY_MODULE_BODY_BASELINE.with(|copy| copy.get()) {
        nc.to_mut();
    }
    // Set centrally for fallback normalizers too. Matching declarations keep
    // the shared view; an actual metadata change obtains a private copy.
    if nc.declared_kind != modifiers.kind { nc.to_mut().declared_kind = modifiers.kind; }
    // Synthesized augmentation bases must be emitted before the using class.
    for link_name in &nc.synthesized_bases {
        let link_key = compiler.canon(link_name);
        let Some(link) = declarations.get(&link_key) else {
            return Err(format!(
                "internal: synthesized base `{link_name}` of class `{cname}` is missing"
            ));
        };
        compiler.compile_class(link)?;
    }
    compiler.compile_class(&nc)
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
                storage,
                ..
            } => {
                let field = NormalField {
                    span: span.clone(),
                    name: name.clone(),
                    type_hint: type_hint.clone(),
                    init: init.clone(),
                    array_bounds: array_bounds.clone(),
                    access: access_from_visibility(modifiers.visibility),
                    readonly: modifiers.is_readonly,
                    value_type: None,
                    // CARRIED, never re-defaulted. This is the one path from a
                    // declared width to anything that can use it; writing
                    // `None` here would drop every `PIC` on the floor and the
                    // only symptom would be records silently landing at
                    // offset 0.
                    storage: *storage,
                };
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
            _ => raw_extra_members.push(member.clone()),
        }
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

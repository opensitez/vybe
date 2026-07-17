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
//! `compiler/classes.rs`. Phase 2b.2 eliminates that reconstruction
//! by porting each pass of `compile_class` to read `NormalClass`
//! fields directly.

use crate::ast::{ClassMember, ClassModifiers, Span, StmtKind};
use crate::compiler::Compiler;
use crate::compiler::class_normalize::{access_from_visibility, from_method_stmt};
use vybe_plugin::class_normalize::types::*;

/// Entry point from `compile_stmt`. Receives the raw AST fields from
/// `StmtKind::ClassDecl`, normalises per language, then hands off to
/// `emit_class`.
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
    // Fill ALL direct bases centrally from the AST parents so no per-language
    // normalizer has to. `parent` stays `parents.first()`; `bases[1..]` is only
    // read behind the `class_multiple_inheritance` opt-in.
    nc.bases = parents.to_vec();
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
                    readonly: modifiers.is_readonly,
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
        bases: Vec::new(),
        interfaces: interfaces.to_vec(),
        is_abstract: modifiers.is_abstract,
        is_sealed: modifiers.is_sealed,
        is_partial: modifiers.is_partial,
        is_value_type: false,
        explicit_self_param: true,
        implicit_self_fields: false,
        instance_fields,
        static_fields,
        instance_methods,
        static_methods,
        properties: Vec::new(),
        constructors: Vec::new(),
        constructor: None,
        destructor: None,
        auto_init_methods: Vec::new(),
        special_methods: Vec::new(),
        event_bindings: Vec::new(),
        raw_extra_members,
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
    vybe_plugin::registry::normalize_class(
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

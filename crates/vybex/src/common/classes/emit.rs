//! `emit_class` — compile a `NormalClass` to bytecode + register its
//! runtime `ClassType` in the `TypeRegistry`.
//!
//! This function is intentionally language-neutral. Every source-language
//! quirk was resolved by the walker's `normalize_class` pass before
//! reaching here. Read `types.rs` and `canonical.rs` for the data
//! contract; `classnormalization.md` for the full architecture.
//!
//! # Phase status (Phase 2b.1 — adapter shim)
//!
//! Currently `emit_class` reconstructs an AST `ClassMember[]` from
//! `NormalClass` and delegates to the legacy `Compiler::compile_class`
//! orchestration. This is a **deliberate no-op port**: the new path
//! is wired end-to-end (walker → normalize_class → emit_class →
//! legacy compile_class), but emits byte-for-byte equivalent bytecode.
//! Proves the plumbing works; pinpoints regressions to the wiring
//! rather than the port.
//!
//! Phase 2b.2 replaces the adapter body with a direct orchestration
//! against `emitter/classes.rs` primitives, one concern at a time
//! (fields → methods → ctor → properties → specials). Each step keeps
//! JS test results steady or improving.

use super::types::*;
use crate::ast::{
    Argument, ClassMember, ClassModifiers, Expression, ExprKind, Modifiers, Param, PassBy,
    PropertySetter, Span, Statement, StmtKind, Visibility,
};
use crate::compiler::Compiler;
use crate::languages::js;

/// Entry point from `compile_stmt`. Receives the raw AST fields from
/// `StmtKind::ClassDecl`, normalises per language, then hands off to
/// `emit_class`. Keeping the AST-to-NormalClass dispatch here (not in
/// the compiler) concentrates language dispatch in one file.
pub fn emit_class_from_ast(
    compiler: &mut Compiler,
    span: Span,
    cname: &str,
    parents: &[String],
    members: &[ClassMember],
    modifiers: &ClassModifiers,
) -> Result<(), String> {
    let nc = normalize_for_profile(
        compiler.profile.name.as_str(),
        span,
        cname,
        parents,
        &[],       // interfaces not yet plumbed from AST
        members,
        modifiers,
    )?;
    emit_class(compiler, nc)
}

/// Dispatch `normalize_class` to the right language implementation based
/// on the profile name. Each language opts in independently as it
/// grows a normalizer; until then, the profile-level
/// `uses_normalize_class` flag stays `false` and this function is
/// unreachable for that language.
fn normalize_for_profile(
    lang: &str,
    span: Span,
    name: &str,
    parents: &[String],
    interfaces: &[String],
    members: &[ClassMember],
    modifiers: &ClassModifiers,
) -> Result<NormalClass, String> {
    match lang {
        "js" => Ok(js::normalize_class::normalize_class(
            span, name, parents, interfaces, members, modifiers,
        )),
        "python" => Ok(crate::languages::python::normalize_class::normalize_class(
            span, name, parents, interfaces, members, modifiers,
        )),
        "ruby" => Ok(crate::languages::ruby::normalize_class::normalize_class(
            span, name, parents, interfaces, members, modifiers,
        )),
        "php" => Ok(crate::languages::php::normalize_class::normalize_class(
            span, name, parents, interfaces, members, modifiers,
        )),
        "dart" => Ok(crate::languages::dart::normalize_class::normalize_class(
            span, name, parents, interfaces, members, modifiers,
        )),
        "pascal" => Ok(crate::languages::pascal::normalize_class::normalize_class(
            span, name, parents, interfaces, members, modifiers,
        )),
        "vb" => Ok(crate::languages::vb::normalize_class::normalize_class(
            span, name, parents, interfaces, members, modifiers,
        )),
        "csharp" => Ok(crate::languages::csharp::normalize_class::normalize_class(
            span, name, parents, interfaces, members, modifiers,
        )),
        other => Err(format!(
            "normalize_class not yet implemented for language {:?} — set \
             `uses_normalize_class = false` in the profile until Phase 3 \
             covers it",
            other
        )),
    }
}

/// Compile a `NormalClass` — the compiler-neutral entry point.
///
/// Phase 2b.1 implementation: reconstruct the equivalent AST
/// `ClassMember[]` and hand off to the legacy `Compiler::compile_class`.
/// This is a shim — no NormalClass-specific emission happens yet.
pub fn emit_class(compiler: &mut Compiler, class: NormalClass) -> Result<(), String> {
    let cname = compiler.canon(&class.name);
    let parent_canonical = class.parent.as_ref().map(|p| compiler.canon(p));
    let members = reconstruct_members(&class);
    compiler.compile_class(&cname, &parent_canonical, &members)
}

/// Rebuild a `Vec<ClassMember>` from a `NormalClass` so the legacy
/// `compile_class` orchestration can consume it unchanged. Used only
/// by the Phase 2b.1 shim — Phase 2b.2 removes this in favour of
/// direct emission.
fn reconstruct_members(class: &NormalClass) -> Vec<ClassMember> {
    let mut members: Vec<ClassMember> = Vec::new();

    // Constructor — single entry; named constructors (Dart) aren't
    // a JS concept and aren't exercised on the pilot path.
    if let Some(ctor) = &class.constructor {
        let base_args = match &ctor.base_call {
            BaseCall::Explicit(args) => Some(args.iter().map(|a| a.value.clone()).collect()),
            BaseCall::Auto | BaseCall::None => None,
        };
        members.push(ClassMember::Constructor {
            params: ctor.params.clone(),
            body: ctor.body.clone(),
            base_args,
            visibility: Visibility::Public,
        });
    }

    // Instance fields.
    for f in &class.instance_fields {
        members.push(ClassMember::Field {
            name: f.name.clone(),
            type_hint: None,
            init: f.init.clone(),
            modifiers: Modifiers { is_static: false, ..Default::default() },
            with_events: false,
            array_bounds: None,
        });
    }

    // Static fields — flagged on Modifiers.
    for f in &class.static_fields {
        members.push(ClassMember::Field {
            name: f.name.clone(),
            type_hint: None,
            init: f.init.clone(),
            modifiers: Modifiers { is_static: true, ..Default::default() },
            with_events: false,
            array_bounds: None,
        });
    }

    // Instance methods — emit each with its `source_name` (not canonical)
    // so the legacy pipeline produces the same bytecode it always did.
    // The canonical-name bookkeeping is Phase 2b.2 territory.
    for m in &class.instance_methods {
        members.push(ClassMember::Method(Box::new(method_to_stmt(m, false))));
    }
    for m in &class.static_methods {
        members.push(ClassMember::Method(Box::new(method_to_stmt(m, true))));
    }

    // Events / Consts / NestedTypes the normalizer didn't explicitly
    // model — forward verbatim so legacy `compile_class` still sees them.
    for m in &class.raw_extra_members {
        members.push(m.clone());
    }

    // Properties.
    for p in &class.properties {
        let getter_body = p.getter.as_ref().map(|g| g.body.clone());
        let setter = p.setter.as_ref().map(|s| PropertySetter {
            param: s.params.first().cloned().unwrap_or_else(|| Param {
                name: "value".into(),
                type_hint: None,
                default: None,
                pass_by: PassBy::Value,
                is_rest: false,
                is_kwargs: false,
                is_optional: false,
                is_nullable: false,
            }),
            body: s.body.clone(),
        });
        members.push(ClassMember::Property {
            name: p.source_name.clone(),
            type_hint: None,
            getter: getter_body,
            setter,
            is_auto: p.auto_field.is_some(),
            modifiers: Modifiers::default(),
        });
    }

    members
}

fn method_to_stmt(m: &NormalMethod, is_static: bool) -> Statement {
    // Start from the walker's raw modifiers to preserve every
    // language-specific flag (is_readonly, is_shared, is_extension,
    // is_overloads, is_not_overridable, decorators). Then stamp the
    // canonical NormalMethod fields on top so they authoritatively
    // win if they ever diverge from the raw struct.
    let mut modifiers = m.raw_modifiers.clone();
    modifiers.is_static = is_static;
    modifiers.is_abstract = m.is_abstract;
    modifiers.is_override = m.is_override;
    modifiers.is_virtual = m.is_virtual;
    modifiers.visibility = access_to_visibility(m.access);

    Statement::new(StmtKind::FunctionDecl {
        name: m.source_name.clone(),
        params: m.params.clone(),
        return_type: m.return_type.clone(),
        body: m.body.clone(),
        modifiers,
        handles: vec![],
        is_async: m.is_async,
        is_generator: m.is_generator,
        is_sub: m.is_sub,
    })
}

/// Invert `access_from_visibility` for reconstruction paths.
fn access_to_visibility(a: Access) -> Visibility {
    match a {
        Access::Public => Visibility::Public,
        Access::Protected => Visibility::Protected,
        Access::Private => Visibility::Private,
        Access::Internal => Visibility::Internal,
    }
}

// Silence rust-unused warnings on AST types that this module re-exports
// implicitly through the reconstruction path.
const _USED: fn() -> Option<ExprKind> = || None;

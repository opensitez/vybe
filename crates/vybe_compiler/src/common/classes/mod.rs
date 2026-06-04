//! Class normalisation — compile-time IR for class declarations,
//! shared across every source language's walker.
//!
//! This module is the compiler side of the class story. The runtime
//! side lives in `vybe_bytecode::component_model::ClassType` (the
//! Component Model wire format). The split:
//!
//!   walker (language-specific)
//!       ↓  produces
//!   NormalClass  ← THIS MODULE
//!       ↓  consumed by
//!   emit::emit_class  ← compiles to bytecode
//!       ↓  registers
//!   vybe_bytecode::ClassType  ← runtime wire format, crosses modules
//!
//! `NormalClass` carries everything the compiler needs (spans, source
//! names, special-method tags, auto-init lists, event bindings, raw
//! AST bodies) that `ClassType` does NOT carry, because those concerns
//! are compile-time-only.
//!
//! See `classnormalization.md` at the project root for the full plan.

pub mod canonical;
pub mod emit;
pub mod types;

pub use canonical::canonicalize_method;
pub use emit::emit_class;
pub use types::{
    Access, BaseCall, EventBinding, NormalClass, NormalConstructor, NormalField, NormalMethod,
    NormalProperty, SpecialMethod, SpecialMethodKind,
};

use crate::ast::{Modifiers, Param, Span, Statement, StmtKind, Visibility};

/// Shared `NormalMethod` builder. Each language's walker calls this
/// after resolving canonical + special-kind info, passing the raw
/// `StmtKind::FunctionDecl` fields plus the language-specific
/// `Access` (derived from `modifiers.visibility` or language rules).
///
/// This preserves every `Modifiers` field verbatim via `raw_modifiers`
/// so reconstruction through the Phase 2b.1 shim is lossless.
pub fn build_normal_method(
    span: Span,
    canonical_name: &str,
    source_name: &str,
    aliases: Vec<String>,
    params: Vec<Param>,
    return_type: Option<String>,
    body: Vec<Statement>,
    access: Access,
    is_async: bool,
    is_generator: bool,
    is_sub: bool,
    raw_modifiers: Modifiers,
) -> NormalMethod {
    NormalMethod {
        span,
        canonical_name: canonical_name.to_string(),
        source_name: source_name.to_string(),
        aliases,
        params,
        return_type,
        body,
        access,
        is_virtual: raw_modifiers.is_virtual,
        is_override: raw_modifiers.is_override,
        is_async,
        is_generator,
        is_abstract: raw_modifiers.is_abstract,
        is_sub,
        raw_modifiers,
    }
}

/// Convenience: extract the `Statement`-shaped `ClassMember::Method`
/// body into a `NormalMethod`, with canonical-name mapping applied.
/// Called from every walker's normalize_class.
pub fn from_method_stmt(
    span: Span,
    stmt: &Statement,
    canonical_name: &str,
    access: Access,
) -> Option<NormalMethod> {
    let StmtKind::FunctionDecl {
        name: src_name,
        params,
        return_type,
        body,
        modifiers,
        is_async,
        is_generator,
        is_sub,
        ..
    } = &stmt.kind
    else {
        return None;
    };
    Some(build_normal_method(
        span,
        canonical_name,
        src_name,
        Vec::new(),
        params.clone(),
        return_type.clone(),
        body.clone(),
        access,
        *is_async,
        *is_generator,
        *is_sub,
        modifiers.clone(),
    ))
}

/// Map a source-level `Visibility` to `Access`. Pascal / C# / VB /
/// Ruby / PHP all use the same enum on the AST side.
pub fn access_from_visibility(v: Visibility) -> Access {
    match v {
        Visibility::Public => Access::Public,
        Visibility::Protected => Access::Protected,
        Visibility::Private => Access::Private,
        Visibility::Internal => Access::Internal,
    }
}

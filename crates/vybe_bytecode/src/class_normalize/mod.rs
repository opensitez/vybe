//! Class normalisation — the language-agnostic compile-time IR for class
//! declarations, produced by every source language's walker.
//!
//!   walker (language-specific)  →  NormalClass (THIS crate)  →
//!   vybe_compiler `class_normalize::emit::emit_class` → bytecode
//!
//! The emit half needs the `Compiler`, so it stays in `vybe_compiler`; the IR
//! + shared builders live here so any language crate can produce them.

pub mod canonical;
pub mod types;

pub use canonical::canonicalize_method;
pub use types::{
    Access, Augmentation, AugmentationAdjustment, AugmentationConflict, AugmentationContributes,
    AugmentationMode, AugmentationPosition, AugmentationSuper, BaseCall, EventBinding, NormalClass,
    NormalConstructor, NormalField, NormalMethod, NormalProperty, SpecialMethod, SpecialMethodKind,
};

use vybe_ast::{Modifiers, Param, Span, Statement, StmtKind, Visibility};

/// Shared `NormalMethod` builder. Each language's walker calls this after
/// resolving canonical + special-kind info, passing the raw
/// `StmtKind::FunctionDecl` fields plus the language-specific `Access`.
/// Preserves every `Modifiers` field verbatim via `raw_modifiers`.
#[allow(clippy::too_many_arguments)]
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

/// Convenience: extract the `Statement`-shaped `ClassMember::Method` body into a
/// `NormalMethod`, with canonical-name mapping applied. Called from every
/// walker's `normalize_class`.
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

/// Map a source-level `Visibility` to `Access`.
pub fn access_from_visibility(v: Visibility) -> Access {
    match v {
        Visibility::Public => Access::Public,
        Visibility::Protected => Access::Protected,
        Visibility::Private => Access::Private,
        Visibility::Internal => Access::Internal,
    }
}

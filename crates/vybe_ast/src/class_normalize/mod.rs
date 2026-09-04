//! Class normalisation — the language-agnostic compile-time IR for class
//! declarations, produced by every source language's walker.
//!
//!   walker (language-specific)  →  NormalClass (THIS crate)  →
//!   vybe_compiler `class_normalize::emit::emit_class` → bytecode
//!
//! The emit half needs the `Compiler`, so it stays in `vybe_compiler`; the IR
//! + shared builders live here so any language crate can produce them.

pub mod types;

pub use types::{
    Access, Augmentation, AugmentationAdjustment, AugmentationConflict, AugmentationContributes,
    AugmentationMode, AugmentationPolicy, AugmentationPosition, AugmentationSuper, BaseCall,
    NormalClass, NormalConstructor, NormalField, NormalMembers, NormalMethod, NormalProperty,
    PlatformBaseSpec, PlatformFieldGui, ProtocolSlot, SpecialMethod, SpecialMethodKind,
    PROTOCOL_SLOT_TABLE,
};

use crate::{
    Argument, ClassMember, ConstructorInitializerTarget, Expression, Modifiers, Param, Span,
    Statement, StmtKind, Visibility,
};

/// Shared `NormalMethod` builder. Each language's walker calls this after
/// resolving canonical + special-kind info, passing the raw
/// `StmtKind::FunctionDecl` fields plus the language-specific `Access`.
/// Preserves every `Modifiers` field verbatim via `raw_modifiers`.
#[allow(clippy::too_many_arguments)]
pub fn build_normal_method(
    span: Span,
    canonical_name: &str,
    source_name: &str,
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

/// The `BaseCall` a constructor performs: what the source SAID, or — when it
/// said nothing and the class has a parent — the implicit parameterless super
/// call.
///
/// `BaseCall::Auto` is a REQUEST, not a commitment: `compile_class` consults
/// `body_has_super_call` and injects nothing when the body already opens with
/// an explicit one. That matters for a language whose walker runs per source
/// file, where the part declaring the constructor may not be the part carrying
/// the inheritance clause.
pub fn base_call_for(
    base_args: Option<&[Expression]>,
    target: ConstructorInitializerTarget,
    has_parent: bool,
) -> BaseCall {
    let Some(args) = base_args else {
        return if has_parent {
            BaseCall::Auto
        } else {
            BaseCall::None
        };
    };
    let args = args.iter().cloned().map(Argument::positional).collect();
    match target {
        ConstructorInitializerTarget::Base => BaseCall::Explicit(args),
        ConstructorInitializerTarget::This => BaseCall::This(args),
    }
}

/// Shared `NormalConstructor` builder for the `ClassMember::Constructor` shape.
///
/// The mapping is mechanical: the member's own fields plus "does this class
/// have a parent" fully determine the result, with no language policy in it.
/// Nine normalizers had written it out identically — the same duplication
/// `normalize_class_from_ast` already calls out for the whole-class fields
/// (`parent`, `is_abstract`, `interfaces`), and the same place to silently
/// disagree, since a language could map `ConstructorInitializerTarget::This`
/// onto `BaseCall::Explicit` and nothing would catch it.
///
/// `has_parent` is the caller's to answer: a normalizer sees the DECLARED
/// parents, while the resolved base list is filled later and centrally.
pub fn from_constructor_member(
    span: Span,
    member: &ClassMember,
    has_parent: bool,
) -> Option<NormalConstructor> {
    let ClassMember::Constructor {
        name,
        params,
        body,
        base_args,
        initializer_target,
        ..
    } = member
    else {
        return None;
    };
    Some(NormalConstructor {
        span,
        params: params.clone(),
        body: body.clone(),
        base_call: base_call_for(base_args.as_deref(), *initializer_target, has_parent),
        named_name: name.clone(),
    })
}

/// The protocol slots this class body DECLARES, as opposed to ones guessed from
/// a method's spelling. Collected before the member loop so a declaration
/// ANYWHERE in the body outranks a guess anywhere else.
pub fn declared_protocol_slots(members: &[ClassMember]) -> Vec<ProtocolSlot> {
    members
        .iter()
        .filter_map(|member| match member {
            ClassMember::Method(stmt) => match &stmt.kind {
                StmtKind::FunctionDecl { modifiers, .. } => modifiers.protocol_slot,
                _ => None,
            },
            _ => None,
        })
        .collect()
}

/// Resolve which special-method slot a method fills: what the walker DECLARED
/// beats what the name table GUESSED from the spelling, and a guess that
/// duplicates some OTHER member's declaration yields entirely.
///
/// A language can spell one slot two ways and mean different things by them.
/// C# `operator ==` defines `a == b` while `Equals` is the virtual
/// object-equality method; VB.NET has exactly the same pair (`Operator =` and
/// `Equals`). Both spellings map to `Eq` through the name table, so a class
/// declaring both collides on one slot key and the loser is silently
/// overwritten — which stopped `a.Equals(b)` from running its own body.
///
/// This is not a per-language rule: evidence outranks inference wherever both
/// spellings exist, so it is answered once here rather than in whichever
/// normalizer happens to hit the bug first.
pub fn resolve_special_kind(
    declared: Option<SpecialMethodKind>,
    guessed: Option<SpecialMethodKind>,
    declared_slots: &[ProtocolSlot],
) -> Option<SpecialMethodKind> {
    declared.or_else(|| guessed.filter(|slot| !declared_slots.contains(slot)))
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

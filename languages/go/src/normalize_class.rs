//! Go `StructDecl` -> `NormalClass` shim.
//!
//! Go was written before class normalization existed, so its structs went
//! through the generic AST fallback and came out with no roles at all — the
//! reason `String()` could not be reached from another language even though
//! the type demonstrably has a text representation.
//!
//! Embedding is declared here as an `Augmentation` with `mode: Promote` — the
//! mode that exists precisely for Go, where a promoted member runs on the
//! INNER value and the receiver rebinds to it. The shared pass carries Go's
//! rule that a shallower depth wins outright and an EQUAL depth with the same
//! name is an error, not a silent pick.

use vybe_ast::{ClassMember, ClassModifiers, Span, StmtKind};
use vybe_bytecode::class_normalize::{
    Access, Augmentation, AugmentationConflict, AugmentationContributes, AugmentationMode,
    AugmentationPosition, AugmentationSuper, NormalClass, NormalField, NormalMembers,
    SpecialMethod, access_from_visibility, from_method_stmt,
};

/// Go field promotion, stated once.
///
/// `AfterOwn` — a method declared on the outer type shadows a promoted one.
/// `Error` — two fields promoting the same name at the same depth is an
/// ambiguous selector, which Go rejects at compile time; picking one silently
/// would compile a program `go build` refuses.
/// Constructors never promote: a Go struct has no constructor member at all.
fn go_embedding(field_name: &str, field_type: &str) -> Augmentation {
    Augmentation {
        from: embedded_type_name(field_type),
        via_field: Some(field_name.to_string()),
        mode: AugmentationMode::Promote,
        position: AugmentationPosition::AfterOwn,
        conflict: AugmentationConflict::Ambiguous,
        super_target: AugmentationSuper::OwnParent,
        adjustments: Vec::new(),
        contributes: AugmentationContributes {
            methods: true,
            // Fields promote in Go too, but a promoted FIELD is a path
            // (`outer.Inner.X`), not a copy — the walker still resolves those.
            // Copying them here would give the outer struct its own storage and
            // silently desynchronise the two.
            fields: false,
            statics: false,
            constructors: false,
            abstract_members: false,
        },
        depth: 0,
    }
}

/// `*pkg.Inner` -> `Inner`. A Go embedded field's NAME is the last segment of
/// its type, which is also how the field is detected as embedded at all.
fn embedded_type_name(type_name: &str) -> String {
    let trimmed = type_name.trim().trim_start_matches('*').trim();
    trimmed
        .rsplit('.')
        .next()
        .unwrap_or(trimmed)
        .trim()
        .to_string()
}

pub fn normalize_class(
    span: Span,
    _name: &str,
    _parents: &[String],
    _interfaces: &[String],
    members: &[ClassMember],
    _modifiers: &ClassModifiers,
) -> NormalClass {
    let mut m = NormalMembers::default();

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
                // A Go embedded field is written with NO name — `struct { Base }`
                // — so the walker gives it the last segment of its type as the
                // name. Field name == type's last segment IS the embedding
                // test; there is no other marker on the AST.
                if let Some(ty) = type_hint.as_deref() {
                    if embedded_type_name(ty) == *name {
                        m.augmentations.push(go_embedding(name, ty));
                    }
                }
                let field = NormalField {
                    span: span.clone(),
                    name: name.clone(),
                    type_hint: type_hint.clone(),
                    init: init.clone(),
                    array_bounds: array_bounds.clone(),
                    access: access_from_visibility(modifiers.visibility),
                    readonly: modifiers.is_readonly,
                };
                m.push_field(modifiers.is_shared, field);
            }
            ClassMember::Method(stmt) => {
                let StmtKind::FunctionDecl {
                    name: source_name,
                    modifiers,
                    ..
                } = &stmt.kind
                else {
                    m.raw_extra_members.push(member.clone());
                    continue;
                };
                let (canonical, special_kind) = crate::protocol::canonical_method(source_name);
                let Some(method) = from_method_stmt(
                    stmt.span.clone(),
                    stmt,
                    &canonical,
                    access_from_visibility(modifiers.visibility),
                ) else {
                    continue;
                };
                if let Some(kind) = special_kind {
                    m.special_methods.push(SpecialMethod {
                        kind,
                        canonical_name: canonical.clone(),
                        source_name: source_name.clone(),
                    });
                }
                m.push_method(modifiers.is_shared, method);
            }
            other => m.raw_extra_members.push(other.clone()),
        }
    }

    NormalClass {
        // A Go method takes its receiver as a declared parameter
        // (`func (s Shape) Area()`), never an implicit `this`.
        explicit_self_param: true,
        ..Default::default()
    }
    .with_members(m)
}

#[cfg(test)]
mod tests {
    use super::*;
    use vybe_ast::{Modifiers, Statement, Visibility};
    use vybe_bytecode::class_normalize::types::SpecialMethodKind;

    fn dummy_span() -> Span {
        Span::default()
    }

    fn make_method(name: &str) -> ClassMember {
        ClassMember::Method(Box::new(Statement::new(StmtKind::FunctionDecl {
            name: name.to_string(),
            params: vec![],
            return_type: None,
            body: vec![],
            modifiers: Modifiers {
                visibility: Visibility::Public,
                ..Default::default()
            },
            handles: vec![],
            is_async: false,
            is_generator: false,
            is_sub: false,
        })))
    }

    /// A Go type fills a role by method NAME — `fmt.Stringer` is `String()
    /// string` and nothing else. Before this normalizer existed the struct
    /// reached the shared model with an empty `special_methods`, so no other
    /// language could resolve its text representation.
    #[test]
    fn stringer_fills_the_tostring_slot() {
        let nc = normalize_class(
            dummy_span(),
            "Point",
            &[],
            &[],
            &[make_method("String")],
            &ClassModifiers::default(),
        );
        assert_eq!(nc.special_methods.len(), 1);
        assert_eq!(nc.special_methods[0].kind, SpecialMethodKind::ToString);
        assert_eq!(nc.special_methods[0].source_name, "String");
    }

    /// `String()` and `Error()` are different interfaces and a type may
    /// declare both, so `Error` must NOT claim the ToString slot — one slot
    /// cannot hold two methods.
    #[test]
    fn error_method_does_not_claim_the_tostring_slot() {
        let nc = normalize_class(
            dummy_span(),
            "MyErr",
            &[],
            &[],
            &[make_method("String"), make_method("Error")],
            &ClassModifiers::default(),
        );
        let tostring: Vec<_> = nc
            .special_methods
            .iter()
            .filter(|s| s.kind == SpecialMethodKind::ToString)
            .collect();
        assert_eq!(tostring.len(), 1);
        assert_eq!(tostring[0].source_name, "String");
        assert_eq!(nc.instance_methods.len(), 2);
    }

    /// An ordinary Go method keeps its own name and claims nothing.
    #[test]
    fn ordinary_method_claims_no_slot() {
        let nc = normalize_class(
            dummy_span(),
            "Point",
            &[],
            &[],
            &[make_method("Area")],
            &ClassModifiers::default(),
        );
        assert!(nc.special_methods.is_empty());
        assert_eq!(nc.instance_methods[0].source_name, "Area");
    }
}

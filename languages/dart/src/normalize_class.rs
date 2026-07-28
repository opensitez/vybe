//! Dart `ClassDecl` → `NormalClass` walker pass.
//!
//! Dart specifics:
//!   - Primary constructor: `ClassName(args) { ... }` or short-form
//!     `ClassName(this.x, this.y);`. Walker produces
//!     `ClassMember::Constructor`.
//!   - Named constructors: `ClassName.named(args)` — walker should
//!     wrap these as methods tagged "constructor.<name>" or similar;
//!     we don't yet distinguish them here (future improvement).
//!   - Factory constructors: `factory ClassName(args) => ...` — static
//!     method that returns an instance; walker represents as a static
//!     method.
//!   - `operator +` / `operator ==` / `operator []` / `operator []=`
//!     → SpecialMethodKind::{Add, Eq, GetItem, SetItem}.
//!   - `get foo => expr` / `set foo(val) { ... }` → NormalProperty.
//!   - `extends A` → parent; `implements I1, I2` → interfaces;
//!     `with M1, M2` (mixins) — walker flattens mixin methods into
//!     member list before we see it.
//!   - `abstract class` → `is_abstract`.
//!   - Dart doesn't have destructors (uses Finalizer from dart:ffi).

use vybe_ast::{ClassMember, ClassModifiers, Modifiers, PropertySetter, Span, StmtKind};
use vybe_bytecode::class_normalize::{
    NormalMembers, build_normal_method,
    from_method_stmt,
    types::*,
};

/// Dart `class X with M` — the language's rules, stated once.
///
/// `Copy`: mixin members are applied into the class. `AfterOwn`: the class's own
/// members beat every mixin. `LastWins`: linearization is left-to-right, so a
/// later mixin overrides an earlier one. `NextInOrder`: `super` inside a mixin
/// method resolves to the next entry in the LINEARIZATION — not the mixin's own
/// parent — which is why Dart needs a real chain and not a flat copy
/// (flexclassplan.md §4c-R). Mixins declare no constructors.
const DART_MIXIN: AugmentationPolicy = AugmentationPolicy {
    mode: AugmentationMode::Copy,
    position: AugmentationPosition::AfterOwn,
    conflict: AugmentationConflict::LastWins,
    super_target: AugmentationSuper::NextInOrder,
    contributes: AugmentationContributes {
        methods: true,
        fields: true,
        statics: false,
        constructors: false,
        abstract_members: true,
    },
};

pub fn normalize_class(
    span: Span,
    name: &str,
    parents: &[String],
    interfaces: &[String],
    members: &[ClassMember],
    modifiers: &ClassModifiers,
) -> NormalClass {
    let mut m = NormalMembers::default();

    for member in members {
        match member {
            ClassMember::Field {
                name: fname,
                type_hint,
                init,
                modifiers: field_modifiers,
                array_bounds,
                ..
            } => {
                let field = NormalField {
                    span: span.clone(),
                    name: fname.clone(),
                    type_hint: type_hint.clone(),
                    init: init.clone(),
                    array_bounds: array_bounds.clone(),
                    access: Access::Public, // Dart's `_name` convention isn't enforced
                    readonly: field_modifiers.is_readonly,
                };
                m.push_field(field_modifiers.is_static, field);
            }
            ClassMember::Method(stmt) => {
                let StmtKind::FunctionDecl {
                    name: src_name,
                    modifiers: method_modifiers,
                    ..
                } = &stmt.kind
                else {
                    continue;
                };
                let (canonical, special_kind) = crate::protocol::canonical_method(src_name);
                let Some(method) = from_method_stmt(span.clone(), stmt, &canonical, Access::Public)
                else {
                    continue;
                };
                if let Some(kind) = special_kind {
                    m.special_methods.push(SpecialMethod {
                        kind,
                        canonical_name: canonical,
                        source_name: src_name.clone(),
                    });
                }
                m.push_method(method_modifiers.is_static, method);
            }
            ClassMember::Constructor {
                name: ctor_name,
                params,
                body,
                base_args,
                ..
            } => {
                let normal = NormalConstructor {
                    span: span.clone(),
                    params: params.clone(),
                    body: body.clone(),
                    base_call: match base_args {
                        Some(args) => BaseCall::Explicit(
                            args.iter()
                                .map(|e| vybe_ast::Argument::positional(e.clone()))
                                .collect(),
                        ),
                        // A Dart subclass ctor without an explicit
                        // `: super(...)` implicitly calls the parent's no-arg
                        // ctor — which is exactly `BaseCall::Auto`, the same
                        // preamble C# and VB already use. A class with no
                        // parent has nothing to call.
                        None if !parents.is_empty() => BaseCall::Auto,
                        None => BaseCall::None,
                    },
                    named_name: ctor_name.clone(),
                };
                // A class can declare an unnamed ctor AND several named ones
                // (`Point(this.x)` + `Point.origin()`); they are distinct
                // constructors, not overloads, and often share an arity. Every
                // one is a variant, and the unnamed one is the primary — which
                // `push_constructor` reads off `named_name`.
                m.push_constructor(normal);
            }
            ClassMember::Property {
                name: pname,
                getter,
                setter,
                is_auto,
                modifiers: prop_modifiers,
                ..
            } => {
                let (canonical, _) = crate::protocol::canonical_method(pname);
                let getter_method = getter.as_ref().map(|body| {
                    build_normal_method(
                        span.clone(),
                        &canonical,
                        pname,
                        vec![],
                        None,
                        body.clone(),
                        Access::Public,
                        false,
                        false,
                        false,
                        Modifiers::default(),
                    )
                });
                let setter_method = setter.as_ref().map(|s: &PropertySetter| {
                    build_normal_method(
                        span.clone(),
                        &canonical,
                        pname,
                        vec![s.param.clone()],
                        None,
                        s.body.clone(),
                        Access::Public,
                        false,
                        false,
                        false,
                        Modifiers::default(),
                    )
                });
                m.properties.push(NormalProperty {
                    span: span.clone(),
                    canonical_name: canonical,
                    source_name: pname.clone(),
                    is_static: prop_modifiers.is_static,
                    getter: getter_method,
                    setter: setter_method,
                    auto_field: if *is_auto { Some(pname.clone()) } else { None },
                });
            }
            // Dart's `with M` is a HEADER clause, not a body member, so it is
            // read from the class header below rather than arriving here.
            ClassMember::Augment(decl) => m.push_augment_decl(decl, DART_MIXIN),
            other @ (ClassMember::Event { .. }
            | ClassMember::Const { .. }
            | ClassMember::NestedType(_)) => {
                m.raw_extra_members.push(other.clone());
            }
        }
    }

    NormalClass {
        // Dart `class X with M` — declared as DATA for the shared
        // augmentation pass rather than folded here. Copy mode: members are
        // duplicated in. `AfterOwn`: the class's own members win. `LastWins`:
        // Dart linearization is left-to-right with later mixins overriding
        // earlier. `NextInOrder`: `super` inside a mixin resolves to the next
        // entry in the linearization, NOT the mixin's own parent. Mixins
        // declare no constructors. See flexclassplan.md §4c.
        augmentations: crate::walker::dart_class_mixins(name)
            .iter()
            .map(|mixin| {
                DART_MIXIN.applied_to(&vybe_ast::AugmentDecl {
                    from: mixin.clone(),
                    ..Default::default()
                })
            })
            .collect(),
        // Dart resolves bare identifiers in instance methods to `this.field`
        // (the `this.` is only needed when a local shadows). Setting this true
        // lets `String toString() => '($x, $y)'` and `_balance += v` reach
        // the instance fields without explicit `this.` qualification.
        implicit_self_fields: true,
        ..Default::default()
    }
    .with_members(m)
}

#[cfg(test)]
mod tests {
    use super::*;
    use vybe_ast::Modifiers;

    fn dummy_span() -> Span {
        Span::default()
    }

    fn make_method(src_name: &str) -> ClassMember {
        ClassMember::Method(Box::new(vybe_ast::Statement::new(StmtKind::FunctionDecl {
            name: src_name.into(),
            params: vec![],
            return_type: None,
            body: vec![],
            modifiers: Modifiers::default(),
            handles: vec![],
            is_async: false,
            is_generator: false,
            is_sub: false,
        })))
    }

    #[test]
    fn to_string_canonicalises_to_tostring() {
        let nc = normalize_class(
            dummy_span(),
            "Foo",
            &[],
            &[],
            &[make_method("toString")],
            &ClassModifiers::default(),
        );
        assert_eq!(nc.instance_methods[0].canonical_name, "tostring");
        assert_eq!(nc.special_methods[0].kind, SpecialMethodKind::ToString);
    }

    #[test]
    fn operator_plus_maps_to_add() {
        let nc = normalize_class(
            dummy_span(),
            "Vec",
            &[],
            &[],
            &[make_method("operator+")],
            &ClassModifiers::default(),
        );
        assert_eq!(nc.instance_methods[0].canonical_name, "add");
        assert_eq!(nc.special_methods[0].kind, SpecialMethodKind::Add);
    }

    #[test]
    fn hash_code_maps_to_hash() {
        let nc = normalize_class(
            dummy_span(),
            "Foo",
            &[],
            &[],
            &[make_method("hashCode")],
            &ClassModifiers::default(),
        );
        assert_eq!(nc.instance_methods[0].canonical_name, "hash");
        assert_eq!(nc.special_methods[0].kind, SpecialMethodKind::Hash);
    }
}

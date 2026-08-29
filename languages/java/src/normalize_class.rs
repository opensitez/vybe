//! Java `ClassDecl` → `NormalClass` normalisation pass.
//!
//! Java is "Java over JS" — the walker produces a JS-shaped common AST
//! and this pass normalises class members into the shared `NormalClass`
//! structure consumed by `common::classes::emit_class`.
//!
//! Java-specific handling:
//!   - Visibility from modifiers (public/protected/private, default = package)
//!   - `final` fields → readonly
//!   - `abstract` / `final` class modifiers
//!   - `implements` interfaces
//!   - Implicit `super()` call when subclass ctor omits it
//!   - `toString()`, `equals()`, `hashCode()`, `compareTo()` → canonical names

use vybe_ast::class_normalize::{NormalMembers, from_method_stmt, types::*};
use vybe_ast::{ClassMember, ClassModifiers, ConstructorInitializerTarget, Span, StmtKind};

/// `class C implements I` — I's DEFAULT methods join C, and C's own win.
///
/// Abstract members do not cross: an interface's body-less members are a
/// requirement ON the implementer, not a gift TO it, and copying them would
/// shadow the implementation the class actually declared.
fn java_interface_defaults() -> AugmentationPolicy {
    AugmentationPolicy {
        mode: AugmentationMode::Copy,
        position: AugmentationPosition::AfterOwn,
        conflict: AugmentationConflict::RequireExplicit,
        super_target: AugmentationSuper::OwnParent,
        contributes: AugmentationContributes {
            methods: true,
            fields: false,
            statics: false,
            constructors: false,
            abstract_members: false,
        },
    }
}

pub fn normalize_class(
    span: Span,
    _name: &str,
    parents: &[String],
    _interfaces: &[String],
    members: &[ClassMember],
    _modifiers: &ClassModifiers,
) -> NormalClass {
    let mut out = NormalMembers::default();

    for member in members {
        match member {
            ClassMember::Field {
                name: fname,
                type_hint,
                init,
                modifiers: m,
                array_bounds,
                ..
            } => {
                let field = NormalField {
                    span: span.clone(),
                    name: fname.clone(),
                    type_hint: type_hint.clone(),
                    init: init.clone(),
                    array_bounds: array_bounds.clone(),
                    access: Access::from(m.visibility),
                    readonly: m.is_readonly,
                    value_type: None,
                    storage: None,
                };
                out.push_field(m.is_static, field);
            }
            ClassMember::Method(stmt) => {
                let StmtKind::FunctionDecl {
                    name: src_name,
                    modifiers: m,
                    ..
                } = &stmt.kind
                else {
                    continue;
                };

                let (canonical, special_kind) = crate::protocol::canonical_method(src_name);
                let access = Access::from(m.visibility);
                // ⚠ NO `is_abstract` FILTER HERE, AND NONE IS NEEDED. There
                // used to be a `continue` on it; removing it changed nothing,
                // because java's WALKER never emits a body-less
                // `method_declaration` as a `ClassMember` at all —
                // `--dump-ast` on `abstract class Base { abstract int foo(); }`
                // contains no `foo`. The contract is dropped a layer earlier
                // than this file, so a filter here can never fire.
                let Some(method) = from_method_stmt(span.clone(), stmt, &canonical, access) else {
                    continue;
                };
                if let Some(kind) = special_kind {
                    out.special_methods.push(SpecialMethod {
                        kind,
                        canonical_name: canonical.clone(),
                        source_name: src_name.clone(),
                    });
                }
                out.push_method(m.is_static, method);
            }
            ClassMember::Constructor {
                params,
                body,
                base_args,
                initializer_target,
                ..
            } => {
                let normalized = NormalConstructor {
                    span: span.clone(),
                    params: params.clone(),
                    body: body.clone(),
                    base_call: match base_args {
                        Some(args) => match initializer_target {
                            ConstructorInitializerTarget::Base => BaseCall::Explicit(
                                args.iter()
                                    .map(|e| vybe_ast::Argument::positional(e.clone()))
                                    .collect(),
                            ),
                            ConstructorInitializerTarget::This => BaseCall::This(
                                args.iter()
                                    .map(|e| vybe_ast::Argument::positional(e.clone()))
                                    .collect(),
                            ),
                        },
                        None => {
                            if parents.is_empty() {
                                BaseCall::None
                            } else {
                                BaseCall::Auto
                            }
                        }
                    },
                    named_name: None,
                };
                out.push_constructor(normalized);
            }
            // ⛔ JAVA'S AUGMENTATION IS INTERFACE `default` METHODS, and this
            // was a no-op arm waiting for a node the walker never emitted.
            //
            // ⚠ IT DOES NOT NEED `Chain`. The note here said it did, and the
            // plan's table said `Chain` + `RequireExplicit` — but `Chain` is
            // Ruby `prepend`, where the augmented member WRAPS the class's own
            // and `super` must reach through it. Java has no such rule:
            //
            //   - a class's own method always wins            → `AfterOwn`
            //   - two interfaces supplying one default, with
            //     no override, is a COMPILE ERROR             → `RequireExplicit`
            //   - `Interface.super.m()` is an EXPLICIT qualified call, not
            //     implicit chaining
            //
            // Measured: `AugmentationMode::Chain` has ONE site — its own
            // declaration, zero readers — while `RequireExplicit` has six and
            // already does exactly this. Kotlin picks `Copy` for the same
            // construct.
            ClassMember::Augment(decl) => {
                out.push_augment_decl(decl, java_interface_defaults());
            }
            other @ (ClassMember::Property { .. }
            | ClassMember::Event { .. }
            | ClassMember::Const { .. }
            | ClassMember::NestedType(_)) => {
                out.raw_extra_members.push(other.clone());
            }
        }
    }

    // Java states nothing beyond its members: `this` is implicit, bare
    // identifiers do not resolve to fields, and classes are reference types.
    NormalClass::default().with_members(out)
}

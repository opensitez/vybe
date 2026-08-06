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

use vybe_ast::{
    ClassMember, ClassModifiers, ConstructorInitializerTarget, Span, StmtKind };
use vybe_ast::class_normalize::{NormalMembers, from_method_stmt, types::*};

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
                    value_type: None };
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
                if m.is_abstract {
                    continue;
                }
                let Some(method) = from_method_stmt(span.clone(), stmt, &canonical, access) else {
                    continue;
                };
                if let Some(kind) = special_kind {
                    out.special_methods.push(SpecialMethod {
                        kind,
                        canonical_name: canonical.clone(),
                        source_name: src_name.clone() });
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
                            ) },
                        None => {
                            if parents.is_empty() {
                                BaseCall::None
                            } else {
                                BaseCall::Auto
                            }
                        }
                    },
                    named_name: None };
                out.push_constructor(normalized);
            }
            // Java's augmentation is interface `default` methods, declared
            // through `implements` in the class HEADER rather than as a body
            // member. Migrating them is flexclassplan.md §4c-R step R7, and it
            // needs `Chain` first.
            ClassMember::Augment(_) => {}
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

//! Fortran `ClassDecl` -> `NormalClass` shim.
//!
//! Fortran derived types (`type :: Foo ... end type`) and type-bound
//! procedures walk directly into the common AST as fields and methods.
//! This shim preserves that shape so the shared class compiler can
//! consume Fortran derived types.

use vybe_ast::class_normalize::{
    Access, BaseCall, NormalClass, NormalConstructor, NormalField, NormalMembers, SpecialMethod,
    from_method_stmt,
};
use vybe_ast::{
    Argument, ClassMember, ClassModifiers, ExprKind, Expression, Literal, Span, StmtKind,
};

fn synthesize_fixed_array_init(bounds: &[Expression]) -> Option<Expression> {
    let size = bounds.first()?.clone();
    Some(Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident("Array")),
        args: vec![
            Argument::positional(size),
            Argument::positional(Expression::new(ExprKind::Lit(Literal::Int(0)))),
        ],
        optional: false,
    }))
}

pub fn normalize_class(
    span: Span,
    _name: &str,
    parents: &[String],
    _interfaces: &[String],
    members: &[ClassMember],
    _modifiers: &ClassModifiers,
) -> NormalClass {
    let mut m = NormalMembers::default();

    for member in members {
        match member {
            ClassMember::Field {
                name: field_name,
                type_hint,
                init,
                modifiers: field_modifiers,
                array_bounds,
                ..
            } => {
                let normalized_type_hint = type_hint.as_ref().map(|hint| {
                    if array_bounds.is_some() && !hint.trim_end().ends_with("()") {
                        format!("{}()", hint.trim())
                    } else {
                        hint.clone()
                    }
                });
                let field = NormalField {
                    span: span.clone(),
                    name: field_name.clone(),
                    type_hint: normalized_type_hint,
                    init: init.clone().or_else(|| {
                        array_bounds
                            .as_deref()
                            .and_then(synthesize_fixed_array_init)
                    }),
                    array_bounds: array_bounds.clone(),
                    access: Access::Public,
                    readonly: field_modifiers.is_readonly,
                    value_type: None,
                };
                m.push_field(field_modifiers.is_static, field);
            }
            ClassMember::Method(stmt) => {
                let StmtKind::FunctionDecl {
                    name: source_name,
                    modifiers: method_modifiers,
                    params,
                    body,
                    ..
                } = &stmt.kind
                else {
                    continue;
                };

                let is_constructor = source_name.eq_ignore_ascii_case("new");
                if is_constructor {
                    m.push_constructor(NormalConstructor {
                        span: span.clone(),
                        params: params.clone(),
                        body: body.clone(),
                        base_call: if parents.is_empty() {
                            BaseCall::None
                        } else {
                            BaseCall::Auto
                        },
                        named_name: None,
                    });
                    continue;
                }

                let (canonical_name, special_kind) = crate::protocol::canonical_method(source_name);
                if let Some(method) =
                    from_method_stmt(span.clone(), stmt, &canonical_name, Access::Public)
                {
                    if let Some(kind) = special_kind {
                        m.special_methods.push(SpecialMethod {
                            kind,
                            canonical_name: canonical_name.clone(),
                            source_name: source_name.to_string(),
                        });
                    }
                    m.push_method(method_modifiers.is_static, method);
                }
            }
            // Fortran derived types extend a single parent (`EXTENDS(base)`) and
            // have no trait/mixin mechanism, so the walker never produces this.
            ClassMember::Augment(_) => {}
            other @ (ClassMember::Constructor { .. }
            | ClassMember::Property { .. }
            | ClassMember::Event { .. }
            | ClassMember::Const { .. }
            | ClassMember::NestedType(_)) => {
                m.raw_extra_members.push(other.clone());
            }
        }
    }

    NormalClass {
        is_value_type: true,
        explicit_self_param: true,
        ..Default::default()
    }
    .with_members(m)
}

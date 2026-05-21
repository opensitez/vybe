//! Fortran `ClassDecl` -> `NormalClass` shim.
//!
//! Fortran derived types (`type :: Foo ... end type`) and type-bound
//! procedures walk directly into the common AST as fields and methods.
//! This shim preserves that shape so the shared class compiler can
//! consume Fortran derived types.

use crate::ast::{Argument, ClassMember, ClassModifiers, ExprKind, Expression, Literal, Span, StmtKind};
use crate::common::classes::{from_method_stmt, Access, BaseCall, NormalClass, NormalConstructor, NormalField};

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
    name: &str,
    parents: &[String],
    interfaces: &[String],
    members: &[ClassMember],
    modifiers: &ClassModifiers,
) -> NormalClass {
    let mut raw_extra_members = Vec::new();
    let mut instance_fields = Vec::new();
    let mut static_fields = Vec::new();
    let mut instance_methods = Vec::new();
    let mut static_methods = Vec::new();
    let mut constructor = None;

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
                    init: init.clone().or_else(|| array_bounds.as_deref().and_then(synthesize_fixed_array_init)),
                    array_bounds: array_bounds.clone(),
                    access: Access::Public,
                    readonly: field_modifiers.is_readonly,
                };
                if field_modifiers.is_static {
                    static_fields.push(field);
                } else {
                    instance_fields.push(field);
                }
            }
            ClassMember::Method(stmt) => {
                let StmtKind::FunctionDecl { name: source_name, modifiers: method_modifiers, params, body, .. } = &stmt.kind else {
                    continue;
                };

                let is_constructor = source_name.eq_ignore_ascii_case("new");
                if is_constructor {
                    constructor = Some(NormalConstructor {
                        span: span.clone(),
                        params: params.clone(),
                        body: body.clone(),
                        base_call: if parents.is_empty() { BaseCall::None } else { BaseCall::Auto },
                        named_name: None,
                    });
                    continue;
                }

                let canonical_name = source_name.to_ascii_lowercase();
                if let Some(method) = from_method_stmt(span.clone(), stmt, &canonical_name, Access::Public) {
                    if method_modifiers.is_static {
                        static_methods.push(method);
                    } else {
                        instance_methods.push(method);
                    }
                }
            }
            other @ (ClassMember::Constructor { .. }
                | ClassMember::Property { .. }
                | ClassMember::Event { .. }
                | ClassMember::Const { .. }
                | ClassMember::NestedType(_)) => {
                raw_extra_members.push(other.clone());
            }
        }
    }

    NormalClass {
        span,
        name: name.to_string(),
        parent: parents.first().cloned(),
        interfaces: interfaces.to_vec(),
        is_abstract: modifiers.is_abstract,
        is_sealed: modifiers.is_sealed,
        is_partial: modifiers.is_partial,
        is_value_type: true,
        explicit_self_param: true,
        implicit_self_fields: false,
        instance_fields,
        static_fields,
        instance_methods,
        static_methods,
        properties: Vec::new(),
        constructors: Vec::new(),
        constructor,
        destructor: None,
        auto_init_methods: Vec::new(),
        special_methods: Vec::new(),
        event_bindings: Vec::new(),
        raw_extra_members,
    }
}

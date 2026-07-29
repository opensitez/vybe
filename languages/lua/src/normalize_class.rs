//! Lua `ClassDecl` -> `NormalClass` shim.
//!
//! Lua's primary OOP surface is still table/metatable normalization in
//! `normalize.rs`, but any Lua AST class declarations should enter the same
//! shared class pipeline as JS/PHP/Ruby/Python. Lua-specific names are resolved
//! here; downstream class emission stays language-neutral.

use vybe_ast::{ClassMember, ClassModifiers, Span, StmtKind};
use vybe_ast::class_normalize::{
    NormalMembers,
    access_from_visibility, from_method_stmt, types::*,
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
                name,
                type_hint,
                init,
                modifiers,
                array_bounds,
                ..
            } => {
                let field = NormalField {
                    span: span.clone(),
                    name: name.clone(),
                    type_hint: type_hint.clone(),
                    init: init.clone(),
                    array_bounds: array_bounds.clone(),
                    access: Access::from(modifiers.visibility),
                    readonly: modifiers.is_readonly,
                };
                m.push_field(modifiers.is_static, field);
            }
            ClassMember::Method(stmt) => {
                let StmtKind::FunctionDecl {
                    name: source_name,
                    modifiers,
                    ..
                } = &stmt.kind
                else {
                    continue;
                };
                let (canonical_name, special_kind) = crate::protocol::canonical_method(source_name);
                let Some(method) = from_method_stmt(
                    span.clone(),
                    stmt,
                    &canonical_name,
                    Access::from(modifiers.visibility),
                ) else {
                    continue;
                };
                if let Some(kind) = special_kind {
                    m.special_methods.push(SpecialMethod {
                        kind,
                        canonical_name,
                        source_name: source_name.clone(),
                    });
                }
                m.push_method(modifiers.is_static, method);
            }
            ClassMember::Constructor {
                params,
                body,
                base_args,
                name,
                ..
            } => {
                m.push_constructor(NormalConstructor {
                    span: span.clone(),
                    params: params.clone(),
                    body: body.clone(),
                    base_call: match base_args {
                        Some(args) => BaseCall::Explicit(
                            args.iter()
                                .map(|arg| vybe_ast::Argument::positional(arg.clone()))
                                .collect(),
                        ),
                        None => BaseCall::None,
                    },
                    named_name: name.clone(),
                });
            }
            ClassMember::Property {
                name,
                type_hint,
                getter,
                setter,
                is_auto,
                modifiers,
            } => {
                let (canonical_name, _) = crate::protocol::canonical_method(name);
                m.properties.push(NormalProperty {
                    span: span.clone(),
                    canonical_name,
                    source_name: name.clone(),
                    is_static: modifiers.is_static,
                    getter: getter.as_ref().and_then(|body| {
                        let stmt = vybe_ast::Statement::new(StmtKind::FunctionDecl {
                            name: name.clone(),
                            params: Vec::new(),
                            return_type: type_hint.clone(),
                            body: body.clone(),
                            modifiers: modifiers.clone(),
                            handles: Vec::new(),
                            is_async: false,
                            is_generator: false,
                            is_sub: false,
                        });
                        from_method_stmt(
                            span.clone(),
                            &stmt,
                            name,
                            Access::from(modifiers.visibility),
                        )
                    }),
                    setter: setter.as_ref().and_then(|setter| {
                        let stmt = vybe_ast::Statement::new(StmtKind::FunctionDecl {
                            name: name.clone(),
                            params: vec![setter.param.clone()],
                            return_type: None,
                            body: setter.body.clone(),
                            modifiers: modifiers.clone(),
                            handles: Vec::new(),
                            is_async: false,
                            is_generator: false,
                            is_sub: true,
                        });
                        from_method_stmt(
                            span.clone(),
                            &stmt,
                            name,
                            Access::from(modifiers.visibility),
                        )
                    }),
                    auto_field: if *is_auto { Some(name.clone()) } else { None },
                });
            }
            // Lua composition is metatable assignment at runtime, not a
            // declaration, so the walker never produces this.
            ClassMember::Augment(_) => {}
            other @ (ClassMember::Event { .. }
            | ClassMember::Const { .. }
            | ClassMember::NestedType(_)) => m.raw_extra_members.push(other.clone()),
        }
    }

    NormalClass {
        explicit_self_param: true,
        ..Default::default()
    }
    .with_members(m)
}

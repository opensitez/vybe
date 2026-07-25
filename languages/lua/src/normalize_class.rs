//! Lua `ClassDecl` -> `NormalClass` shim.
//!
//! Lua's primary OOP surface is still table/metatable normalization in
//! `normalize.rs`, but any Lua AST class declarations should enter the same
//! shared class pipeline as JS/PHP/Ruby/Python. Lua-specific names are resolved
//! here; downstream class emission stays language-neutral.

use vybe_ast::{ClassMember, ClassModifiers, Span, StmtKind};
use vybe_bytecode::class_normalize::{
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
    let mut raw_extra_members = Vec::new();
    let mut instance_fields = Vec::new();
    let mut static_fields = Vec::new();
    let mut instance_methods = Vec::new();
    let mut static_methods = Vec::new();
    let mut properties = Vec::new();
    let mut constructor = None;
    let mut special_methods = Vec::new();

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
                    access: access_from_visibility(modifiers.visibility),
                    readonly: modifiers.is_readonly,
                };
                if modifiers.is_static {
                    static_fields.push(field);
                } else {
                    instance_fields.push(field);
                }
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
                    access_from_visibility(modifiers.visibility),
                ) else {
                    continue;
                };
                if let Some(kind) = special_kind {
                    special_methods.push(SpecialMethod {
                        kind,
                        canonical_name,
                        source_name: source_name.clone(),
                    });
                }
                if modifiers.is_static {
                    static_methods.push(method);
                } else {
                    instance_methods.push(method);
                }
            }
            ClassMember::Constructor {
                params,
                body,
                base_args,
                name,
                ..
            } => {
                constructor = Some(NormalConstructor {
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
                properties.push(NormalProperty {
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
                            access_from_visibility(modifiers.visibility),
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
                            access_from_visibility(modifiers.visibility),
                        )
                    }),
                    auto_field: if *is_auto { Some(name.clone()) } else { None },
                });
            }
            other @ (ClassMember::Event { .. }
            | ClassMember::Const { .. }
            | ClassMember::NestedType(_)) => raw_extra_members.push(other.clone()),
        }
    }

    NormalClass {
        span,
        name: name.to_string(),
        parent: parents.first().cloned(),
        bases: Vec::new(),
        interfaces: interfaces.to_vec(),
        is_abstract: modifiers.is_abstract,
        is_sealed: modifiers.is_sealed,
        is_partial: modifiers.is_partial,
        is_value_type: false,
        explicit_self_param: true,
        implicit_self_fields: false,
        instance_fields,
        static_fields,
        instance_methods,
        static_methods,
        properties,
        constructors: Vec::new(),
        constructor,
        destructor: None,
        auto_init_methods: Vec::new(),
        special_methods,
        event_bindings: Vec::new(),
        raw_extra_members,
    }
}

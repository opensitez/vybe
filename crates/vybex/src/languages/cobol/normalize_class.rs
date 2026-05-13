//! COBOL `ClassDecl` -> `NormalClass` shim.
//!
//! COBOL classes currently walk directly into the common AST as fields and
//! methods. This shim preserves that shape so the shared class compiler can
//! consume OO COBOL without needing a dedicated semantic lowering pass yet.

use crate::ast::{ClassMember, ClassModifiers, Span, StmtKind};
use crate::common::classes::{from_method_stmt, Access, BaseCall, NormalClass, NormalConstructor, NormalField};

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
            ClassMember::Field { name: field_name, type_hint, init, modifiers: field_modifiers, .. } => {
                let field = NormalField {
                    span: span.clone(),
                    name: field_name.clone(),
                    type_hint: type_hint.clone(),
                    init: init.clone(),
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
        is_value_type: false,
        explicit_self_param: false,
        implicit_self_fields: true,
        instance_fields,
        static_fields,
        instance_methods,
        static_methods,
        properties: Vec::new(),
        constructor,
        destructor: None,
        auto_init_methods: Vec::new(),
        special_methods: Vec::new(),
        event_bindings: Vec::new(),
        raw_extra_members,
    }
}
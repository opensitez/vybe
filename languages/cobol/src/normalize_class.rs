//! COBOL `ClassDecl` -> `NormalClass` shim.
//!
//! COBOL classes currently walk directly into the common AST as fields and
//! methods. This shim preserves that shape so the shared class compiler can
//! consume OO COBOL without needing a dedicated semantic lowering pass yet.

use vybe_ast::class_normalize::{
    Access, BaseCall, NormalClass, NormalConstructor, NormalField, NormalMembers, from_method_stmt,
};
use vybe_ast::{ClassMember, ClassModifiers, Span, StmtKind};

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
                let field = NormalField {
                    span: span.clone(),
                    name: field_name.clone(),
                    type_hint: type_hint.clone(),
                    init: init.clone(),
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

                let canonical_name = source_name.to_ascii_lowercase();
                if let Some(method) =
                    from_method_stmt(span.clone(), stmt, &canonical_name, Access::Public)
                {
                    m.push_method(method_modifiers.is_static, method);
                }
            }
            // COBOL has no trait/mixin/module mechanism — `INHERITS` is plain
            // single inheritance — so the walker never produces this.
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
        implicit_self_fields: true,
        ..Default::default()
    }
    .with_members(m)
}

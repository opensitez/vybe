//! Kotlin `ClassDecl` → `NormalClass` normalization pass.

use vybe_ast::{
    ClassMember, ClassModifiers, ConstructorInitializerTarget, Span, StmtKind,
};
use vybe_ast::class_normalize::{from_method_stmt, types::*, NormalMembers};

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
            other => {
                out.raw_extra_members.push(other.clone());
            }
        }
    }

    NormalClass {
        implicit_self_fields: true,
        ..Default::default()
    }
    .with_members(out)
}

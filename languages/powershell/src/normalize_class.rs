//! PowerShell class declaration normalizer.
//!
//! The walker already lowered class syntax into `ClassDecl` + `ClassMember`.
//! This normalizer only maps members into the shared `NormalMembers` shape.

use vybe_ast::class_normalize::{NormalMembers, build_normal_method, from_method_stmt, types::*};
use vybe_ast::{
    ClassMember, ClassModifiers, ConstructorInitializerTarget, PropertySetter, Span, StmtKind,
};

pub fn normalize_class(
    span: Span,
    name: &str,
    parents: &[String],
    interfaces: &[String],
    members: &[ClassMember],
    modifiers: &ClassModifiers,
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
                out.push_field(m.is_static || m.is_shared, field);
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
                let Some(method) = from_method_stmt(span.clone(), stmt, &canonical, access) else {
                    continue;
                };

                if special_kind == Some(SpecialMethodKind::Destructor) {
                    out.destructor = Some(method);
                    continue;
                }

                if let Some(kind) = special_kind {
                    out.special_methods.push(SpecialMethod {
                        kind,
                        canonical_name: canonical.clone(),
                        source_name: src_name.clone(),
                    });
                }

                out.push_method(m.is_static || m.is_shared, method);
            }
            ClassMember::Constructor {
                params,
                body,
                base_args,
                initializer_target,
                ..
            } => {
                let base_call = match base_args {
                    Some(args) => match initializer_target {
                        ConstructorInitializerTarget::Base => BaseCall::Explicit(
                            args.iter()
                                .map(|expr| vybe_ast::Argument::positional(expr.clone()))
                                .collect(),
                        ),
                        ConstructorInitializerTarget::This => BaseCall::This(
                            args.iter()
                                .map(|expr| vybe_ast::Argument::positional(expr.clone()))
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
                };

                out.push_constructor(NormalConstructor {
                    span: span.clone(),
                    params: params.clone(),
                    body: body.clone(),
                    base_call,
                    named_name: None,
                });
            }
            // A property reaches here ONLY when the walker built one for a
            // validated field (`[ValidateRange(1,100)][int]$Score`): PowerShell
            // has no `get`/`set` block syntax, so an ordinary member is a
            // `Field`. Both accessors are fully implemented — `auto_field` stays
            // `None`, because the backing field is a separate member under its
            // own name and an auto field would collide with the accessor.
            ClassMember::Property {
                name: pname,
                getter,
                setter,
                modifiers: m,
                ..
            } => {
                let access = Access::from(m.visibility);
                let getter_method = getter.as_ref().map(|body| {
                    build_normal_method(
                        span.clone(),
                        pname,
                        pname,
                        vec![],
                        None,
                        body.clone(),
                        access,
                        false,
                        false,
                        false,
                        m.clone(),
                    )
                });
                let setter_method = setter.as_ref().map(|st: &PropertySetter| {
                    build_normal_method(
                        span.clone(),
                        pname,
                        pname,
                        vec![st.param.clone()],
                        None,
                        st.body.clone(),
                        access,
                        false,
                        false,
                        false,
                        m.clone(),
                    )
                });
                // ⛔The canonical name is LOWERCASED, unlike C#'s, and that is
                // load-bearing. PowerShell is case-insensitive, so member access
                // folds `$b.Score` to `score` — but the accessor chunk is named
                // from `canonical_name`, so passing the declared spelling
                // emitted `__get_Score` while every read looked up
                // `__get_score`. Both accessors existed and neither ever fired:
                // the read fell through to the property's own field slot, which
                // no initializer had written, so a validated property answered
                // empty and accepted any assignment.
                out.properties.push(NormalProperty {
                    span: span.clone(),
                    canonical_name: pname.to_lowercase(),
                    source_name: pname.clone(),
                    is_static: m.is_static,
                    getter: getter_method,
                    setter: setter_method,
                    auto_field: None,
                });
            }
            ClassMember::Event { .. }
            | ClassMember::Const { .. }
            | ClassMember::NestedType(_)
            | ClassMember::Augment(_) => {
                out.raw_extra_members.push(member.clone());
            }
        }
    }

    let mut normalized = NormalClass {
        name: name.to_string(),
        is_partial: modifiers.is_partial,
        is_abstract: modifiers.is_abstract,
        is_sealed: modifiers.is_sealed,
        ..Default::default()
    };

    normalized.parent = parents.first().cloned();
    normalized.bases = parents.to_vec();
    normalized.interfaces = interfaces.to_vec();

    normalized.with_members(out)
}

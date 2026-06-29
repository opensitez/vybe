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

use crate::ast::{
    ClassMember, ClassModifiers, ConstructorInitializerTarget, Span, StmtKind, Visibility,
};
use crate::common::classes::{
    canonical::{ClassLang, canonicalize_method},
    from_method_stmt,
    types::*,
};

pub fn normalize_class(
    span: Span,
    name: &str,
    parents: &[String],
    interfaces: &[String],
    members: &[ClassMember],
    modifiers: &ClassModifiers,
) -> NormalClass {
    let mut raw_extra_members: Vec<ClassMember> = Vec::new();
    let mut instance_fields: Vec<NormalField> = Vec::new();
    let mut static_fields: Vec<NormalField> = Vec::new();
    let mut instance_methods: Vec<NormalMethod> = Vec::new();
    let mut static_methods: Vec<NormalMethod> = Vec::new();
    let mut constructors: Vec<NormalConstructor> = Vec::new();
    let mut constructor: Option<NormalConstructor> = None;
    let mut special_methods: Vec<SpecialMethod> = Vec::new();

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
                    access: access_from_visibility(m.visibility),
                    readonly: m.is_readonly,
                };
                if m.is_static {
                    static_fields.push(field);
                } else {
                    instance_fields.push(field);
                }
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

                let (canonical, special_kind) = canonicalize_method(ClassLang::Java, src_name);
                let access = access_from_visibility(m.visibility);
                let Some(method) = from_method_stmt(span.clone(), stmt, &canonical, access) else {
                    continue;
                };
                if let Some(kind) = special_kind {
                    special_methods.push(SpecialMethod {
                        kind,
                        canonical_name: canonical.clone(),
                        source_name: src_name.clone(),
                    });
                }
                if m.is_static {
                    static_methods.push(method);
                } else {
                    instance_methods.push(method);
                }
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
                                    .map(|e| crate::ast::Argument::positional(e.clone()))
                                    .collect(),
                            ),
                            ConstructorInitializerTarget::This => BaseCall::This(
                                args.iter()
                                    .map(|e| crate::ast::Argument::positional(e.clone()))
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
                constructor = Some(normalized.clone());
                constructors.push(normalized);
            }
            other @ (ClassMember::Property { .. }
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
        is_partial: false,
        is_value_type: false,
        explicit_self_param: false,
        implicit_self_fields: false,
        instance_fields,
        static_fields,
        instance_methods,
        static_methods,
        properties: Vec::new(),
        constructors,
        constructor,
        destructor: None,
        auto_init_methods: Vec::new(),
        special_methods,
        event_bindings: Vec::new(),
        raw_extra_members,
    }
}

fn access_from_visibility(v: Visibility) -> Access {
    match v {
        Visibility::Public => Access::Public,
        Visibility::Protected => Access::Protected,
        Visibility::Private => Access::Private,
        Visibility::Internal => Access::Internal,
    }
}

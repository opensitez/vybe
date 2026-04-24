//! Pascal `ClassDecl` → `NormalClass` walker pass.
//!
//! Pascal / Delphi / Free Pascal class specifics:
//!   - `constructor Create;` / `constructor Init;` → NormalConstructor.
//!     Pascal's convention is `Create`; Free Pascal also allows `Init`.
//!   - `destructor Destroy;` → destructor.
//!   - `property Foo read GetFoo write SetFoo` → NormalProperty. Walker
//!     already links property accessors to their accessor methods.
//!   - `class operator Add(...)` / `class operator Equal(...)` →
//!     SpecialMethodKind::Add / Eq. Pascal operator overloads arrive
//!     with names like "Add" / "Subtract" / "Multiply" / "Divide" /
//!     "Equal" per Delphi convention.
//!   - `override` / `virtual` / `reintroduce` → flag carries through.
//!   - Case-insensitive: Pascal method names lowercase to canonical.

use crate::ast::{ClassMember, ClassModifiers, Modifiers, PropertySetter, Span, Statement, StmtKind};
use crate::common::classes::{
    build_normal_method,
    canonical::{canonicalize_method, ClassLang},
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
    let mut properties: Vec<NormalProperty> = Vec::new();
    let mut constructor: Option<NormalConstructor> = None;
    let mut destructor: Option<NormalMethod> = None;
    let mut special_methods: Vec<SpecialMethod> = Vec::new();

    for member in members {
        match member {
            ClassMember::Field { name: fname, init, modifiers: m, .. } => {
                let field = NormalField {
                    span: span.clone(),
                    name: fname.clone(),
                    init: init.clone(),
                    access: Access::Public,
                    readonly: m.is_readonly,
                };
                if m.is_static {
                    static_fields.push(field);
                } else {
                    instance_fields.push(field);
                }
            }
            ClassMember::Method(stmt) => {
                let StmtKind::FunctionDecl { name: src_name, modifiers: m, .. } = &stmt.kind else {
                    continue;
                };

                // Pascal destructor: `destructor Destroy;`. Case-insensitive.
                if src_name.eq_ignore_ascii_case("Destroy") {
                    if let Some(d) = from_method_stmt(span.clone(), stmt, "destructor", Access::Public) {
                        destructor = Some(d);
                    }
                    continue;
                }

                let (canonical, special_kind) = canonicalize_method(ClassLang::Pascal, src_name);
                let Some(method) = from_method_stmt(span.clone(), stmt, &canonical, Access::Public) else {
                    continue;
                };
                if let Some(kind) = special_kind {
                    special_methods.push(SpecialMethod {
                        kind,
                        canonical_name: canonical,
                        source_name: src_name.clone(),
                    });
                }
                if m.is_static {
                    static_methods.push(method);
                } else {
                    instance_methods.push(method);
                }
            }
            ClassMember::Constructor { params, body, base_args, .. } => {
                constructor = Some(NormalConstructor {
                    span: span.clone(),
                    params: params.clone(),
                    body: body.clone(),
                    base_call: match base_args {
                        Some(args) => BaseCall::Explicit(
                            args.iter().map(|e| crate::ast::Argument::positional(e.clone())).collect(),
                        ),
                        // Pascal: `inherited;` or `inherited Create;` is
                        // the explicit call. Walker today emits base_args
                        // = None when absent — mirror with Auto if there's
                        // a parent, None otherwise.
                        None => if parents.is_empty() { BaseCall::None } else { BaseCall::Auto },
                    },
                    named_name: None,
                });
            }
            ClassMember::Property { name: pname, getter, setter, is_auto, .. } => {
                let (canonical, _) = canonicalize_method(ClassLang::Pascal, pname);
                let getter_method = getter.as_ref().map(|body| build_normal_method(
                    span.clone(), &canonical, pname, Vec::new(),
                    vec![], None, body.clone(),
                    Access::Public, false, false, false, Modifiers::default(),
                ));
                let setter_method = setter.as_ref().map(|s: &PropertySetter| build_normal_method(
                    span.clone(), &canonical, pname, Vec::new(),
                    vec![s.param.clone()], None, s.body.clone(),
                    Access::Public, false, false, false, Modifiers::default(),
                ));
                properties.push(NormalProperty {
                    span: span.clone(),
                    canonical_name: canonical,
                    source_name: pname.clone(),
                    getter: getter_method,
                    setter: setter_method,
                    auto_field: if *is_auto { Some(pname.clone()) } else { None },
                });
            }
            other @ (ClassMember::Event { .. } | ClassMember::Const { .. } | ClassMember::NestedType(_)) => {
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
        instance_fields,
        static_fields,
        instance_methods,
        static_methods,
        properties,
        constructor,
        destructor,
        auto_init_methods: Vec::new(),
        special_methods,
        event_bindings: Vec::new(),
        raw_extra_members,
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::ast::Modifiers;

    fn dummy_span() -> Span { Span::default() }

    fn make_method(src_name: &str) -> ClassMember {
        ClassMember::Method(Box::new(Statement::new(StmtKind::FunctionDecl {
            name: src_name.into(),
            params: vec![],
            return_type: None,
            body: vec![],
            modifiers: Modifiers::default(),
            handles: vec![],
            is_async: false,
            is_generator: false,
            is_sub: false,
        })))
    }

    #[test]
    fn destroy_goes_to_destructor_case_insensitive() {
        let nc = normalize_class(
            dummy_span(), "Foo", &[], &[], &[make_method("Destroy")],
            &ClassModifiers::default(),
        );
        assert!(nc.destructor.is_some());
        assert!(nc.instance_methods.is_empty());

        // case variant
        let nc2 = normalize_class(
            dummy_span(), "Foo", &[], &[], &[make_method("destroy")],
            &ClassModifiers::default(),
        );
        assert!(nc2.destructor.is_some());
    }

    #[test]
    fn add_operator_maps_to_canonical_add() {
        let nc = normalize_class(
            dummy_span(), "Vec", &[], &[], &[make_method("Add")],
            &ClassModifiers::default(),
        );
        assert_eq!(nc.instance_methods[0].canonical_name, "add");
        assert_eq!(nc.special_methods[0].kind, SpecialMethodKind::Add);
    }
}

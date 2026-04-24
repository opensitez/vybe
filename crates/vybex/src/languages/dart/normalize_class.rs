//! Dart `ClassDecl` → `NormalClass` walker pass.
//!
//! Dart specifics:
//!   - Primary constructor: `ClassName(args) { ... }` or short-form
//!     `ClassName(this.x, this.y);`. Walker produces
//!     `ClassMember::Constructor`.
//!   - Named constructors: `ClassName.named(args)` — walker should
//!     wrap these as methods tagged "constructor.<name>" or similar;
//!     we don't yet distinguish them here (future improvement).
//!   - Factory constructors: `factory ClassName(args) => ...` — static
//!     method that returns an instance; walker represents as a static
//!     method.
//!   - `operator +` / `operator ==` / `operator []` / `operator []=`
//!     → SpecialMethodKind::{Add, Eq, GetItem, SetItem}.
//!   - `get foo => expr` / `set foo(val) { ... }` → NormalProperty.
//!   - `extends A` → parent; `implements I1, I2` → interfaces;
//!     `with M1, M2` (mixins) — walker flattens mixin methods into
//!     member list before we see it.
//!   - `abstract class` → `is_abstract`.
//!   - Dart doesn't have destructors (uses Finalizer from dart:ffi).

use crate::ast::{ClassMember, ClassModifiers, Modifiers, PropertySetter, Span, StmtKind};
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
    let mut special_methods: Vec<SpecialMethod> = Vec::new();

    for member in members {
        match member {
            ClassMember::Field { name: fname, init, modifiers: m, .. } => {
                let field = NormalField {
                    span: span.clone(),
                    name: fname.clone(),
                    init: init.clone(),
                    access: Access::Public, // Dart's `_name` convention isn't enforced
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
                let (canonical, special_kind) = canonicalize_method(ClassLang::Dart, src_name);
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
                        // Dart: subclass ctor without explicit `: super(...)`
                        // auto-calls the no-arg super ctor in real Dart, but
                        // the Vybe compiler doesn't yet implement that
                        // auto-insertion. Walker mirrors the source — emit
                        // layer can opt in later by switching this to Auto.
                        None => BaseCall::None,
                    },
                    named_name: None, // TODO: plumb Dart named ctors when walker marks them
                });
            }
            ClassMember::Property { name: pname, getter, setter, is_auto, .. } => {
                let (canonical, _) = canonicalize_method(ClassLang::Dart, pname);
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
        explicit_self_param: false,
        implicit_self_fields: false, // Dart requires `this.field` for ambiguous refs
        instance_fields,
        static_fields,
        instance_methods,
        static_methods,
        properties,
        constructor,
        destructor: None,
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
    fn toString_canonicalises_to_tostring() {
        let nc = normalize_class(
            dummy_span(), "Foo", &[], &[], &[make_method("toString")],
            &ClassModifiers::default(),
        );
        assert_eq!(nc.instance_methods[0].canonical_name, "tostring");
        assert_eq!(nc.special_methods[0].kind, SpecialMethodKind::ToString);
    }

    #[test]
    fn operator_plus_maps_to_add() {
        let nc = normalize_class(
            dummy_span(), "Vec", &[], &[], &[make_method("operator+")],
            &ClassModifiers::default(),
        );
        assert_eq!(nc.instance_methods[0].canonical_name, "add");
        assert_eq!(nc.special_methods[0].kind, SpecialMethodKind::Add);
    }

    #[test]
    fn hashCode_maps_to_hash() {
        let nc = normalize_class(
            dummy_span(), "Foo", &[], &[], &[make_method("hashCode")],
            &ClassModifiers::default(),
        );
        assert_eq!(nc.instance_methods[0].canonical_name, "hash");
        assert_eq!(nc.special_methods[0].kind, SpecialMethodKind::Hash);
    }
}

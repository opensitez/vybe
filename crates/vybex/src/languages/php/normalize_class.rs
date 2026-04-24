//! PHP `ClassDecl` → `NormalClass` walker pass.
//!
//! Per-language responsibilities:
//!   - `__construct` → `NormalConstructor`.
//!   - `__destruct` → `destructor`.
//!   - `__toString` → canonical `tostring` (SpecialMethodKind::ToString).
//!   - `__invoke` / `__call` → canonical `call` (SpecialMethodKind::Call).
//!   - `__get` / `__set` / `__unset` → attribute-access interceptors
//!     (SpecialMethodKind::GetAttr/SetAttr/DelAttr).
//!   - Visibility keywords `public` / `protected` / `private` from the
//!     method/field `modifiers.visibility`.
//!   - `readonly` fields carried through.
//!   - Traits (`use Trait;`) are flattened by the walker BEFORE reaching
//!     here — the member list contains the class's own members plus
//!     every trait method copied in (PHP's semantics).
//!   - `abstract class` / `final class` → `is_abstract` / `is_sealed`.
//!   - `class Foo implements I1, I2` → `interfaces`.

use crate::ast::{ClassMember, ClassModifiers, Modifiers, PropertySetter, Span, StmtKind, Visibility};
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
                let StmtKind::FunctionDecl { name: src_name, modifiers: m, .. } = &stmt.kind else {
                    continue;
                };

                // __destruct gets routed away from methods list.
                if src_name == "__destruct" {
                    if let Some(d) = from_method_stmt(
                        span.clone(), stmt, "destructor",
                        access_from_visibility(m.visibility),
                    ) {
                        destructor = Some(d);
                    }
                    continue;
                }

                let (canonical, special_kind) = canonicalize_method(ClassLang::Php, src_name);
                let access = access_from_visibility(m.visibility);
                let Some(method) = from_method_stmt(span.clone(), stmt, &canonical, access) else {
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
                        // PHP: `parent::__construct(...)` is optional;
                        // when missing no auto-call is emitted.
                        None => BaseCall::None,
                    },
                    named_name: None,
                });
            }
            ClassMember::Property { name: pname, getter, setter, is_auto, modifiers: m, .. } => {
                let (canonical, _) = canonicalize_method(ClassLang::Php, pname);
                let access = access_from_visibility(m.visibility);
                let getter_method = getter.as_ref().map(|body| build_normal_method(
                    span.clone(), &canonical, pname, Vec::new(),
                    vec![], None, body.clone(),
                    access, false, false, false, Modifiers::default(),
                ));
                let setter_method = setter.as_ref().map(|s: &PropertySetter| build_normal_method(
                    span.clone(), &canonical, pname, Vec::new(),
                    vec![s.param.clone()], None, s.body.clone(),
                    access, false, false, false, Modifiers::default(),
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
        implicit_self_fields: false, // PHP requires `$this->field`
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

fn access_from_visibility(v: Visibility) -> Access {
    match v {
        Visibility::Public => Access::Public,
        Visibility::Protected => Access::Protected,
        Visibility::Private => Access::Private,
        Visibility::Internal => Access::Internal,
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
    fn tostring_magic_maps_to_canonical() {
        let nc = normalize_class(
            dummy_span(), "Foo", &[], &[], &[make_method("__toString")],
            &ClassModifiers::default(),
        );
        assert_eq!(nc.instance_methods[0].canonical_name, "tostring");
        assert_eq!(nc.special_methods[0].kind, SpecialMethodKind::ToString);
    }

    #[test]
    fn destruct_goes_to_destructor_slot() {
        let nc = normalize_class(
            dummy_span(), "Foo", &[], &[], &[make_method("__destruct")],
            &ClassModifiers::default(),
        );
        assert!(nc.destructor.is_some());
        assert!(nc.instance_methods.is_empty());
    }

    #[test]
    fn invoke_and_call_both_map_to_call_kind() {
        let nc_invoke = normalize_class(
            dummy_span(), "Foo", &[], &[], &[make_method("__invoke")],
            &ClassModifiers::default(),
        );
        assert_eq!(nc_invoke.special_methods[0].kind, SpecialMethodKind::Call);

        let nc_call = normalize_class(
            dummy_span(), "Foo", &[], &[], &[make_method("__call")],
            &ClassModifiers::default(),
        );
        assert_eq!(nc_call.special_methods[0].kind, SpecialMethodKind::Call);
    }

    #[test]
    fn get_and_set_map_to_getattr_setattr() {
        let nc = normalize_class(
            dummy_span(), "Foo", &[], &[],
            &[make_method("__get"), make_method("__set")],
            &ClassModifiers::default(),
        );
        let kinds: Vec<_> = nc.special_methods.iter().map(|s| s.kind).collect();
        assert!(kinds.contains(&SpecialMethodKind::GetAttr));
        assert!(kinds.contains(&SpecialMethodKind::SetAttr));
    }

    #[test]
    fn interfaces_passed_through() {
        let nc = normalize_class(
            dummy_span(), "Foo", &[], &["Countable".into(), "IteratorAggregate".into()],
            &[], &ClassModifiers::default(),
        );
        assert_eq!(nc.interfaces, vec!["Countable".to_string(), "IteratorAggregate".to_string()]);
    }
}

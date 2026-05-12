//! Python `ClassDecl` → `NormalClass` walker pass.
//!
//! The Python walker (`walker.rs::walk_class_def` +
//! `stmts_to_class_members`) has already done the first pass:
//! `__init__` is wrapped as `ClassMember::Constructor`, class-level
//! assignments are static fields, and everything else is a Method.
//!
//! This pass does the cross-language normalisation:
//!   - Canonicalises special method names (`__str__` → `tostring`,
//!     `__add__` → `add`, …) via the shared canonical-name table.
//!   - Stamps `SpecialMethodKind` when a method is operator / protocol
//!     overloadable.
//!   - Detects `@staticmethod` / `@classmethod` / `@property` from the
//!     walker's existing first-param-name heuristic (walker already
//!     flips `modifiers.is_static` for methods whose first param isn't
//!     `self`). Property decorators aren't yet captured — future
//!     improvement in the Python walker, not here.
//!   - Maps visibility: Python uses a leading underscore as *convention*
//!     for "protected" and double-leading for "mangled private", but
//!     these are not enforced; treat as `Public` and let the compiler
//!     emit whatever the walker produced (consistent with Python's
//!     own "consenting adults" philosophy).

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
    _interfaces: &[String],
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
            ClassMember::Field { name: fname, type_hint, init, modifiers: m, .. } => {
                let field = NormalField {
                    span: span.clone(),
                    name: fname.clone(),
                    type_hint: type_hint.clone(),
                    init: init.clone(),
                    access: Access::Public, // Python is convention-based
                    readonly: false,
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

                // Python `__del__` is the finaliser — route to destructor.
                if src_name == "__del__" {
                    if let Some(d) = from_method_stmt(span.clone(), stmt, src_name, Access::Public) {
                        destructor = Some(d);
                    }
                    continue;
                }

                let (canonical, special_kind) = canonicalize_method(ClassLang::Python, src_name);
                let Some(method) = from_method_stmt(span.clone(), stmt, &canonical, Access::Public) else {
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
            ClassMember::Constructor { params, body, base_args, .. } => {
                constructor = Some(NormalConstructor {
                    span: span.clone(),
                    params: params.clone(),
                    body: body.clone(),
                    base_call: match base_args {
                        Some(args) => BaseCall::Explicit(
                            args.iter().map(|e| crate::ast::Argument::positional(e.clone())).collect(),
                        ),
                        // Python: `super().__init__()` is the conventional
                        // explicit call; when absent, no auto-call is
                        // injected (Python's object.__init__ takes no args
                        // and is a no-op, so subclasses are free to skip
                        // calling it — MRO finds it if needed).
                        None => BaseCall::None,
                    },
                    named_name: None,
                });
            }
            ClassMember::Property { name: pname, getter, setter, is_auto, modifiers: m, .. } => {
                // Python `@property` / `@foo.setter` aren't yet captured
                // by the walker — if this arm fires at all today, it's
                // from a dunder mapping in the walker. Keep it general.
                let (canonical, _) = canonicalize_method(ClassLang::Python, pname);
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
                    is_static: m.is_static,
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
        // Python multiple inheritance — walker currently puts all parents
        // in `parents`. The first becomes the principal superclass; any
        // remaining go into `interfaces` so `isinstance` can still check
        // them. C3 linearisation of methods is NOT done here yet — future
        // work when emit_class is direct and can flatten mixed-in methods.
        interfaces: if parents.len() > 1 { parents[1..].to_vec() } else { Vec::new() },
        is_abstract: modifiers.is_abstract,
        is_sealed: modifiers.is_sealed,
        is_partial: false,
        explicit_self_param: true,
        implicit_self_fields: false,
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
    use crate::ast::{Expression, ExprKind, Literal, Modifiers};

    fn dummy_span() -> Span { Span::default() }

    #[test]
    fn str_method_maps_to_canonical_tostring() {
        let method = crate::ast::Statement::new(StmtKind::FunctionDecl {
            name: "__str__".into(),
            params: vec![crate::ast::Param {
                name: "self".into(), type_hint: None, default: None,
                pass_by: crate::ast::PassBy::Value, is_rest: false,
                is_kwargs: false, is_optional: false, is_nullable: false,
            }],
            return_type: None,
            body: vec![],
            modifiers: Modifiers::default(),
            handles: vec![],
            is_async: false,
            is_generator: false,
            is_sub: false,
        });
        let members = vec![ClassMember::Method(Box::new(method))];
        let nc = normalize_class(
            dummy_span(), "Foo", &[], &[], &members, &ClassModifiers::default(),
        );
        assert_eq!(nc.instance_methods.len(), 1);
        assert_eq!(nc.instance_methods[0].canonical_name, "tostring");
        assert_eq!(nc.instance_methods[0].source_name, "__str__");
        assert_eq!(nc.special_methods.len(), 1);
        assert_eq!(nc.special_methods[0].kind, SpecialMethodKind::ToString);
    }

    #[test]
    fn add_operator_special_method() {
        let method = crate::ast::Statement::new(StmtKind::FunctionDecl {
            name: "__add__".into(),
            params: vec![],
            return_type: None,
            body: vec![],
            modifiers: Modifiers::default(),
            handles: vec![],
            is_async: false,
            is_generator: false,
            is_sub: false,
        });
        let members = vec![ClassMember::Method(Box::new(method))];
        let nc = normalize_class(
            dummy_span(), "Vec2", &[], &[], &members, &ClassModifiers::default(),
        );
        assert_eq!(nc.instance_methods[0].canonical_name, "add");
        assert_eq!(nc.special_methods[0].kind, SpecialMethodKind::Add);
    }

    #[test]
    fn del_becomes_destructor() {
        let method = crate::ast::Statement::new(StmtKind::FunctionDecl {
            name: "__del__".into(),
            params: vec![],
            return_type: None,
            body: vec![],
            modifiers: Modifiers::default(),
            handles: vec![],
            is_async: false,
            is_generator: false,
            is_sub: false,
        });
        let members = vec![ClassMember::Method(Box::new(method))];
        let nc = normalize_class(
            dummy_span(), "Foo", &[], &[], &members, &ClassModifiers::default(),
        );
        assert!(nc.destructor.is_some());
        assert!(nc.instance_methods.is_empty());
    }

    #[test]
    fn multiple_inheritance_first_is_parent_rest_are_interfaces() {
        let nc = normalize_class(
            dummy_span(),
            "Child",
            &["A".to_string(), "B".to_string(), "C".to_string()],
            &[], &[], &ClassModifiers::default(),
        );
        assert_eq!(nc.parent.as_deref(), Some("A"));
        assert_eq!(nc.interfaces, vec!["B".to_string(), "C".to_string()]);
    }

    #[test]
    fn constructor_without_super_call_is_none_basecall() {
        let member = ClassMember::Constructor {
            params: vec![],
            body: vec![],
            base_args: None,
            visibility: crate::ast::Visibility::Public,
        };
        let nc = normalize_class(
            dummy_span(), "Foo", &[], &[], &[member], &ClassModifiers::default(),
        );
        // Python: no auto-super-call; skipping is legal.
        assert!(matches!(nc.constructor.as_ref().unwrap().base_call, BaseCall::None));
    }

    #[test]
    fn static_field_lands_in_static_fields_list() {
        let init = Expression::new(ExprKind::Lit(Literal::Int(42)));
        let mut mods = Modifiers::default();
        mods.is_static = true;
        let member = ClassMember::Field {
            name: "COUNT".into(),
            type_hint: None,
            init: Some(init),
            modifiers: mods,
            with_events: false,
            array_bounds: None,
        };
        let nc = normalize_class(
            dummy_span(), "Foo", &[], &[], &[member], &ClassModifiers::default(),
        );
        assert_eq!(nc.static_fields.len(), 1);
        assert!(nc.instance_fields.is_empty());
    }
}

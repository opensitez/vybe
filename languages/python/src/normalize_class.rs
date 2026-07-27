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

use vybe_ast::{ClassMember, ClassModifiers, Modifiers, PropertySetter, Span, StmtKind};
use vybe_bytecode::class_normalize::{
    NormalMembers,
    build_normal_method,
    canonical::{ClassLang, canonicalize_method},
    from_method_stmt,
    types::*,
};

/// `<ClassName>.<field>` — the read an instance-side mirror of a class
/// attribute initialises from, so the value object is shared rather than
/// re-constructed per instance.
fn class_attr_read(class_name: &str, field: &str) -> vybe_ast::Expression {
    vybe_ast::Expression::new(vybe_ast::ExprKind::Member {
        object: Box::new(vybe_ast::Expression::new(vybe_ast::ExprKind::Ident(
            class_name.to_string(),
        ))),
        field: field.to_string(),
        null_safe: false,
    })
}

pub fn normalize_class(
    span: Span,
    name: &str,
    parents: &[String],
    _interfaces: &[String],
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
                    access: Access::Public, // Python is convention-based
                    readonly: false,
                };
                // A Python class attribute is readable through instances
                // (`a.kind` falls back to `type(a).kind`), so the class body's
                // `kind = ...` is BOTH a static field and an instance one. The
                // shared router owns the doubling; Python supplies only what is
                // Python-specific — how to name the class attribute to read
                // from.
                if m.is_static {
                    out.push_static_field_readable_on_instances(
                        field,
                        class_attr_read(name, fname),
                    );
                } else {
                    out.push_field(false, field);
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

                let (canonical, special_kind) = canonicalize_method(ClassLang::Python, src_name);
                let Some(method) = from_method_stmt(span.clone(), stmt, &canonical, Access::Public)
                else {
                    continue;
                };

                // `__del__` is the finaliser — a lifecycle member, not a
                // method. Routed by KIND; the spelling is declared once in the
                // shared canonical table.
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

                out.push_method(m.is_static, method);
            }
            ClassMember::Constructor {
                params,
                body,
                base_args,
                ..
            } => {
                out.push_constructor(NormalConstructor {
                    span: span.clone(),
                    params: params.clone(),
                    body: body.clone(),
                    base_call: match base_args {
                        Some(args) => BaseCall::Explicit(
                            args.iter()
                                .map(|e| vybe_ast::Argument::positional(e.clone()))
                                .collect(),
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
            ClassMember::Property {
                name: pname,
                getter,
                setter,
                is_auto,
                modifiers: m,
                ..
            } => {
                // Python `@property` / `@foo.setter` aren't yet captured
                // by the walker — if this arm fires at all today, it's
                // from a dunder mapping in the walker. Keep it general.
                let (canonical, _) = canonicalize_method(ClassLang::Python, pname);
                let getter_method = getter.as_ref().map(|body| {
                    build_normal_method(
                        span.clone(),
                        &canonical,
                        pname,
                        Vec::new(),
                        vec![],
                        None,
                        body.clone(),
                        Access::Public,
                        false,
                        false,
                        false,
                        Modifiers::default(),
                    )
                });
                let setter_method = setter.as_ref().map(|s: &PropertySetter| {
                    build_normal_method(
                        span.clone(),
                        &canonical,
                        pname,
                        Vec::new(),
                        vec![s.param.clone()],
                        None,
                        s.body.clone(),
                        Access::Public,
                        false,
                        false,
                        false,
                        Modifiers::default(),
                    )
                });
                out.properties.push(NormalProperty {
                    span: span.clone(),
                    canonical_name: canonical,
                    source_name: pname.clone(),
                    is_static: m.is_static,
                    getter: getter_method,
                    setter: setter_method,
                    auto_field: if *is_auto { Some(pname.clone()) } else { None },
                });
            }
            // Python has no separate augmentation syntax — a "mixin" is just
            // another base class, so it arrives through `parents` and is
            // resolved by the C3 MRO, not by this pass.
            ClassMember::Augment(_) => {}
            other @ (ClassMember::Event { .. }
            | ClassMember::Const { .. }
            | ClassMember::NestedType(_)) => {
                out.raw_extra_members.push(other.clone());
            }
        }
    }

    NormalClass {
        // Python multiple inheritance — walker currently puts all parents
        // in `parents`. The first becomes the principal superclass; any
        // remaining go into `interfaces` so `isinstance` can still check
        // them. C3 linearisation of methods is NOT done here yet — future
        // work when emit_class is direct and can flatten mixed-in methods.
        // Python multiple inheritance: bases beyond the first join the
        // interface list so `isinstance` answers for all of them. An ADDITION
        // to the declared interfaces, which are filled centrally.
        interfaces: if parents.len() > 1 {
            parents[1..].to_vec()
        } else {
            Vec::new()
        },
        explicit_self_param: true,
        ..Default::default()
    }
    .with_members(out)
}

#[cfg(test)]
mod tests {
    use super::*;
    use vybe_ast::{ExprKind, Expression, Literal, Modifiers};

    fn dummy_span() -> Span {
        Span::default()
    }

    #[test]
    fn str_method_maps_to_canonical_tostring() {
        let method = vybe_ast::Statement::new(StmtKind::FunctionDecl {
            name: "__str__".into(),
            params: vec![vybe_ast::Param {
                name: "self".into(),
                type_hint: None,
                default: None,
                pass_by: vybe_ast::PassBy::Value,
                is_rest: false,
                is_kwargs: false,
                is_optional: false,
                is_nullable: false,
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
            dummy_span(),
            "Foo",
            &[],
            &[],
            &members,
            &ClassModifiers::default(),
        );
        assert_eq!(nc.instance_methods.len(), 1);
        assert_eq!(nc.instance_methods[0].canonical_name, "tostring");
        assert_eq!(nc.instance_methods[0].source_name, "__str__");
        assert_eq!(nc.special_methods.len(), 1);
        assert_eq!(nc.special_methods[0].kind, SpecialMethodKind::ToString);
    }

    #[test]
    fn add_operator_special_method() {
        let method = vybe_ast::Statement::new(StmtKind::FunctionDecl {
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
            dummy_span(),
            "Vec2",
            &[],
            &[],
            &members,
            &ClassModifiers::default(),
        );
        assert_eq!(nc.instance_methods[0].canonical_name, "add");
        assert_eq!(nc.special_methods[0].kind, SpecialMethodKind::Add);
    }

    #[test]
    fn del_becomes_destructor() {
        let method = vybe_ast::Statement::new(StmtKind::FunctionDecl {
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
            dummy_span(),
            "Foo",
            &[],
            &[],
            &members,
            &ClassModifiers::default(),
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
            &[],
            &[],
            &ClassModifiers::default(),
        );
        assert_eq!(nc.parent.as_deref(), Some("A"));
        assert_eq!(nc.interfaces, vec!["B".to_string(), "C".to_string()]);
    }

    #[test]
    fn constructor_without_super_call_is_none_basecall() {
        let member = ClassMember::Constructor {
            name: None,
            params: vec![],
            body: vec![],
            base_args: None,
            initializer_target: vybe_ast::ConstructorInitializerTarget::Base,
            visibility: vybe_ast::Visibility::Public,
        };
        let nc = normalize_class(
            dummy_span(),
            "Foo",
            &[],
            &[],
            &[member],
            &ClassModifiers::default(),
        );
        // Python: no auto-super-call; skipping is legal.
        assert!(matches!(
            nc.constructor.as_ref().unwrap().base_call,
            BaseCall::None
        ));
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
            dummy_span(),
            "Foo",
            &[],
            &[],
            &[member],
            &ClassModifiers::default(),
        );
        assert_eq!(nc.static_fields.len(), 1);
        assert!(nc.instance_fields.is_empty());
    }
}

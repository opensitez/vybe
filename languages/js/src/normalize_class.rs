//! JS `ClassDecl` → `NormalClass` walker pass.
//!
//! Reads the AST shape the JS grammar produces (`StmtKind::ClassDecl`
//! with `Vec<ClassMember>`) and returns a `NormalClass` suitable for
//! the shared `common::classes::emit_class` pipeline. All JS-specific
//! class concepts are resolved HERE so downstream compilation is
//! language-neutral.
//!
//! Resolves at normalisation time:
//!   - `constructor() { … }` → `NormalConstructor`, with `BaseCall`
//!     inferred from the body (first statement = `super(args)` →
//!     `Explicit`; otherwise `None` for root classes or `Auto` for
//!     subclasses once we support it).
//!   - `static foo() { … }` → `static_methods`.
//!   - `get` / `set foo()` → `NormalProperty`.
//!   - `#private` field / method → `Access::Private`.
//!   - `[Symbol.iterator]()` → canonical name `iterator` +
//!     `SpecialMethodKind::Iterator`. Also `Symbol.{asyncIterator,
//!     toPrimitive, hasInstance}`.
//!   - `toString() { … }` / `valueOf()` — mapped via the JS canonical
//!     table so cross-language consumers still find them under the
//!     right canonical name.
//!   - Class fields (`x = 5` at class body) → `instance_fields` with
//!     `init` expression; `static x = 5` → `static_fields`.
//!
//! The registered frontend normalizer supplies the owned `NormalClass` member
//! data consumed by the common class compiler. Its input is borrowed, so
//! method bodies and parameters must still be copied at this boundary.

use vybe_ast::class_normalize::{NormalMembers, build_normal_method, from_method_stmt, types::*};
use vybe_ast::{
    Argument, ClassMember, ClassModifiers, ExprKind, Expression, LambdaBody, Modifiers, Param,
    PropertySetter, Span, Statement, StmtKind,
};

/// Normalise the members of a JS `class X extends Y { … }` declaration.
///
/// The common pipeline supplies the class header. The source walker resolves
/// a nontrivial `extends` expression into a preceding synthetic binding, so
/// this adapter only needs to normalize the borrowed member declarations.
pub fn normalize_class(
    span: Span,
    _name: &str,
    _parents: &[String],
    _interfaces: &[String], // JS has no interface concept
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
                modifiers,
                array_bounds,
                ..
            } => {
                let field = NormalField {
                    span: span.clone(),
                    name: fname.clone(),
                    type_hint: type_hint.clone(),
                    init: init.clone(),
                    array_bounds: array_bounds.clone(),
                    access: access_for_js(fname),
                    readonly: false, // JS doesn't have readonly at class field level
                    value_type: None,
                    storage: None,
                };
                out.push_field(modifiers.is_static, field);
            }
            ClassMember::Method(stmt) => {
                let StmtKind::FunctionDecl { name, body, modifiers, .. } = &stmt.kind else {
                    continue;
                };
                if modifiers.is_static && (name == "__static_init" || name == "__static_init__") {
                    let block_index = out.static_fields.len();
                    out.push_field(true, static_block_field(span.clone(), block_index, body.clone()));
                    continue;
                }
                let (canonical, kind) = crate::protocol::canonical_method_ref(name);
                if let Some(nm) = from_method_stmt(span.clone(), stmt, canonical, access_for_js(name)) {
                    if let Some(kind) = kind {
                        out.special_methods.push(SpecialMethod {
                            kind,
                            canonical_name: nm.canonical_name.clone(),
                            source_name: nm.source_name.clone(),
                        });
                    }
                    out.push_method(modifiers.is_static, nm);
                }
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
                        // JS: subclass without explicit `super()` is a
                        // runtime TypeError in the spec, but JS's grammar
                        // permits it. The walker mirrors the source
                        // faithfully — no auto super-call insertion.
                        None => BaseCall::None,
                    },
                    named_name: None, // JS has no named constructors
                });
            }
            ClassMember::Property {
                name: pname,
                getter,
                setter,
                is_auto,
                modifiers,
                ..
            } => {
                let (canon, _kind) = crate::protocol::canonical_method_ref(pname);
                let mut prop_mods = Modifiers::default();
                prop_mods.is_override = modifiers.is_override;
                let getter_method = getter.as_ref().map(|body| {
                    build_normal_method(
                        span.clone(),
                        &canon,
                        pname,
                        vec![],
                        None,
                        body.clone(),
                        access_for_js(pname),
                        false,
                        false,
                        false,
                        prop_mods.clone(),
                    )
                });
                let setter_method = setter.as_ref().map(|s: &PropertySetter| {
                    build_normal_method(
                        span.clone(),
                        &canon,
                        pname,
                        vec![s.param.clone()],
                        None,
                        s.body.clone(),
                        access_for_js(pname),
                        false,
                        false,
                        false,
                        prop_mods.clone(),
                    )
                });
                out.properties.push(NormalProperty {
                    raw_setter: None,
                    span: span.clone(),
                    canonical_name: canon.to_owned(),
                    source_name: pname.clone(),
                    is_static: modifiers.is_static,
                    getter: getter_method,
                    setter: setter_method,
                    auto_field: if *is_auto { Some(pname.clone()) } else { None },
                });
            }
            // `Event`, `Const`, `NestedType` aren't reached from JS AST
            // (those are VB / C# / Pascal constructs). Keep the match
            // exhaustive so future ClassMember additions force a
            // conscious choice here.
            // JS has no augmentation declaration: a "mixin" is a plain function
            // that mutates a prototype at runtime, which the prototype fallback
            // (§4d) already handles. Nothing static to declare.
            ClassMember::Augment(_) => {}
            other @ (ClassMember::Event { .. }
            | ClassMember::Const { .. }
            | ClassMember::NestedType(_)) => {
                out.raw_extra_members.push(other.clone());
            }
        }
    }

    NormalClass {
        implicit_self_fields: false, // JS: bare `foo` doesn't resolve to this.foo
        ..Default::default()         // JS has no destructor syntax
    }
    .with_members(out)
}

fn static_block_field(span: Span, index: usize, body: Vec<Statement>) -> NormalField {
    let lambda = Expression::new(ExprKind::Lambda {
        params: Vec::<Param>::new(),
        body: LambdaBody::Block(body),
        is_async: false,
        captures: Vec::new(),
    });
    let init = Expression::new(ExprKind::Call {
        callee: Box::new(lambda),
        args: Vec::<Argument>::new(),
        optional: false,
    });
    NormalField {
        span,
        name: format!("__static_block_{}", index),
        type_hint: None,
        init: Some(init),
        array_bounds: None,
        access: Access::Private,
        readonly: false,
        value_type: None,
        storage: None,
    }
}

/// JS `#private` names get `Access::Private`. Everything else is `Public`.
/// Inner classes / sub-modules aren't a JS thing, so no `Internal`.
fn access_for_js(name: &str) -> Access {
    if name.starts_with('#') {
        Access::Private
    } else {
        Access::Public
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn canonical_resolution_preserves_owned_parameter_body_and_declared_slot() {
        let body = vec![Statement::new(StmtKind::Expr(Expression::string("use strict")))];
        let param = Param {
            name: "value".into(),
            default: Some(Expression::int(7)),
            type_hint: None,
            pass_by: vybe_ast::PassBy::Value,
            is_rest: false,
            is_kwargs: false,
            is_optional: false,
            is_nullable: false,
        };
        let modifiers = Modifiers {
            is_static: true,
            protocol_slot: Some(SpecialMethodKind::Iterator),
            ..Default::default()
        };
        let member = ClassMember::Method(Box::new(Statement::new(StmtKind::FunctionDecl {
            name: "ordinary".into(), params: vec![param.clone()], return_type: None,
            body: body.clone(), modifiers: modifiers.clone(), handles: vec![],
            is_async: true, is_generator: false, is_sub: false,
        })));
        let result = normalize_class(dummy_span(), "C", &[], &[], &[member], &ClassModifiers::default());
        let method = &result.static_methods[0];
        assert_eq!(method.canonical_name, "ordinary");
        assert_eq!(method.source_name, "ordinary");
        assert_eq!(format!("{:?}", method.params), format!("{:?}", vec![param]));
        assert_eq!(format!("{:?}", method.body), format!("{body:?}"));
        assert_eq!(format!("{:?}", method.raw_modifiers), format!("{modifiers:?}"));
        assert!(method.is_async);
    }

    fn dummy_span() -> Span {
        Span::default()
    }

    #[test]
    fn empty_class_produces_empty_normal_class() {
        let nc = normalize_class(
            dummy_span(),
            "Foo",
            &[],
            &[],
            &[],
            &ClassModifiers::default(),
        );
        // The shared normalize_class_from_ast pipeline supplies the class
        // header. This frontend adapter owns member semantics only.
        assert!(nc.instance_methods.is_empty());
        assert!(nc.constructor.is_none());
    }

    #[test]
    fn to_string_method_gets_canonical_name_and_special_kind() {
        use vybe_ast::{Modifiers, Statement, StmtKind};
        let method = Statement::new(StmtKind::FunctionDecl {
            name: "toString".into(),
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
        assert_eq!(nc.instance_methods.len(), 1);
        assert_eq!(nc.instance_methods[0].canonical_name, "tostring");
        assert_eq!(nc.instance_methods[0].source_name, "toString");
        assert_eq!(nc.special_methods.len(), 1);
        assert_eq!(nc.special_methods[0].kind, SpecialMethodKind::ToString);
    }

    #[test]
    fn hash_prefixed_member_is_private() {
        use vybe_ast::{Modifiers, Statement, StmtKind};
        let method = Statement::new(StmtKind::FunctionDecl {
            name: "#secret".into(),
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
        assert_eq!(nc.instance_methods[0].access, Access::Private);
    }

    #[test]
    fn static_method_lands_in_static_list() {
        use vybe_ast::{Modifiers, Statement, StmtKind};
        let method = Statement::new(StmtKind::FunctionDecl {
            name: "create".into(),
            params: vec![],
            return_type: None,
            body: vec![],
            modifiers: Modifiers {
                is_static: true,
                ..Default::default()
            },
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
        assert_eq!(nc.static_methods.len(), 1);
        assert!(nc.instance_methods.is_empty());
    }

    #[test]
    fn constructor_with_super_call_records_explicit_base_call() {
        use vybe_ast::{ExprKind, Expression};
        let base_arg = Expression::new(ExprKind::Ident("x".into()));
        let member = ClassMember::Constructor {
            name: None,
            params: vec![],
            body: vec![],
            base_args: Some(vec![base_arg]),
            initializer_target: vybe_ast::ConstructorInitializerTarget::Base,
            visibility: vybe_ast::Visibility::Public,
        };
        let nc = normalize_class(
            dummy_span(),
            "Dog",
            &["Animal".to_string()],
            &[],
            &[member],
            &ClassModifiers::default(),
        );
        assert!(matches!(
            nc.constructor.as_ref().unwrap().base_call,
            BaseCall::Explicit(_)
        ));
        let BaseCall::Explicit(args) = &nc.constructor.as_ref().unwrap().base_call else {
            unreachable!()
        };
        assert_eq!(args.len(), 1);
        assert!(matches!(&args[0].value.kind, ExprKind::Ident(name) if name == "x"));
    }

    #[test]
    fn subclass_without_explicit_super_gets_none_base_call() {
        // JS spec: derived class missing `super()` is a runtime TypeError.
        // The walker doesn't auto-insert a super call; it faithfully
        // reports `BaseCall::None` so downstream emission treats the
        // missing super as user error, not a normalization default.
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
            "Dog",
            &["Animal".to_string()],
            &[],
            &[member],
            &ClassModifiers::default(),
        );
        assert!(matches!(
            nc.constructor.as_ref().unwrap().base_call,
            BaseCall::None
        ));
    }

    #[test]
    fn root_class_without_super_gets_none_base_call() {
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
            "Animal",
            &[],
            &[],
            &[member],
            &ClassModifiers::default(),
        );
        assert!(matches!(
            nc.constructor.as_ref().unwrap().base_call,
            BaseCall::None
        ));
    }
}

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
//! # Phase 2a status
//!
//! This module is wired up additively — it produces a `NormalClass`
//! but nothing consumes it yet. The JS compile path still goes through
//! the legacy `compile_class` orchestration in `crate::compiler::classes`.
//! Phase 2b flips the switch.

use vybe_ast::{
    Argument, ClassMember, ClassModifiers, ExprKind, Expression, LambdaBody, Modifiers, Param,
    PropertySetter, Span, Statement, StmtKind,
};
use vybe_bytecode::class_normalize::{
    build_normal_method,
    canonical::{ClassLang, canonicalize_method},
    from_method_stmt,
    types::*,
};

/// Normalise the members of a JS `class X extends Y { … }` declaration.
///
/// `parents[0]` is the direct superclass name if any (the JS grammar
/// captures `extends <Expression>` but the walker currently only
/// forwards ident-shaped expressions — computed-super is tracked as a
/// known limitation; see `test_class_patterns::class_extends_expression`).
pub fn normalize_class(
    span: Span,
    name: &str,
    parents: &[String],
    _interfaces: &[String], // JS has no interface concept
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
                };
                if modifiers.is_static {
                    static_fields.push(field);
                } else {
                    instance_fields.push(field);
                }
            }
            ClassMember::Method(stmt) => {
                if let Some(nm) = method_from_funcdecl(span.clone(), stmt) {
                    if is_static_method(stmt)
                        && (nm.source_name == "__static_init"
                            || nm.source_name == "__static_init__")
                    {
                        static_fields.push(static_block_field(
                            span.clone(),
                            static_fields.len(),
                            nm.body,
                        ));
                        continue;
                    }
                    let (canon, kind) = canonicalize_method(ClassLang::Js, &nm.source_name);
                    let nm = NormalMethod {
                        canonical_name: canon.clone(),
                        ..nm
                    };
                    if let Some(k) = kind {
                        special_methods.push(SpecialMethod {
                            kind: k,
                            canonical_name: canon,
                            source_name: nm.source_name.clone(),
                        });
                    }
                    if is_static_method(stmt) {
                        static_methods.push(nm);
                    } else {
                        instance_methods.push(nm);
                    }
                }
            }
            ClassMember::Constructor {
                params,
                body,
                base_args,
                ..
            } => {
                constructor = Some(NormalConstructor {
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
                let (canon, _kind) = canonicalize_method(ClassLang::Js, pname);
                let mut prop_mods = Modifiers::default();
                prop_mods.is_override = modifiers.is_override;
                let getter_method = getter.as_ref().map(|body| {
                    build_normal_method(
                        span.clone(),
                        &canon,
                        pname,
                        Vec::new(),
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
                        Vec::new(),
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
                properties.push(NormalProperty {
                    span: span.clone(),
                    canonical_name: canon,
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
            other @ (ClassMember::Event { .. }
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
        bases: Vec::new(),
        interfaces: Vec::new(),
        is_abstract: modifiers.is_abstract,
        is_sealed: modifiers.is_sealed,
        is_partial: false,
        is_value_type: false,
        explicit_self_param: false,
        implicit_self_fields: false, // JS: bare `foo` doesn't resolve to this.foo
        instance_fields,
        static_fields,
        instance_methods,
        static_methods,
        properties,
        constructors: Vec::new(),
        constructor,
        destructor: None, // JS has no destructor syntax
        auto_init_methods: Vec::new(),
        special_methods,
        event_bindings: Vec::new(),
        raw_extra_members,
    }
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
    }
}

/// Build a `NormalMethod` from a `StmtKind::FunctionDecl` wrapped
/// inside a `ClassMember::Method`. The canonical name starts as the
/// source name; caller overwrites after applying
/// `canonicalize_method`. Returns `None` if the inner kind is
/// anything other than `FunctionDecl`.
fn method_from_funcdecl(span: Span, stmt: &Statement) -> Option<NormalMethod> {
    let StmtKind::FunctionDecl { name, .. } = &stmt.kind else {
        return None;
    };
    let access = access_for_js(name);
    from_method_stmt(span, stmt, name, access)
}

fn is_static_method(stmt: &Statement) -> bool {
    matches!(
        &stmt.kind,
        StmtKind::FunctionDecl { modifiers, .. } if modifiers.is_static
    )
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
        assert_eq!(nc.name, "Foo");
        assert!(nc.parent.is_none());
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
        assert_eq!(nc.parent.as_deref(), Some("Animal"));
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

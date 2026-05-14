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

use crate::ast::{ClassMember, ClassModifiers, Expression, ExprKind, Modifiers, PropertySetter, Span, Statement, StmtKind};
use crate::common::classes::{
    build_normal_method,
    canonical::{canonicalize_method, ClassLang},
    from_method_stmt,
    types::*,
};
use std::collections::{HashMap, HashSet};

fn property_field_name(body: &[Statement], field_names: &HashSet<String>) -> Option<String> {
    let [stmt] = body else { return None; };
    match &stmt.kind {
        StmtKind::Return(Some(expr)) => match &expr.kind {
            ExprKind::Call { callee, args, .. } if args.is_empty() => match &callee.kind {
                ExprKind::Member { object, field, .. } if matches!(object.kind, ExprKind::This) => {
                    field_names.contains(&field.to_ascii_lowercase()).then(|| field.clone())
                }
                _ => None,
            },
            _ => None,
        },
        StmtKind::Expr(expr) => match &expr.kind {
            ExprKind::Call { callee, args, .. } if args.len() == 1 => match &callee.kind {
                ExprKind::Member { object, field, .. }
                    if matches!(object.kind, ExprKind::This)
                        && matches!(args[0].value.kind, ExprKind::Ident(ref name) if name.eq_ignore_ascii_case("value")) =>
                {
                    field_names.contains(&field.to_ascii_lowercase()).then(|| field.clone())
                }
                _ => None,
            },
            _ => None,
        },
        _ => None,
    }
}

fn rewrite_property_getter_body(body: &[Statement], field_names: &HashSet<String>) -> Vec<Statement> {
    let Some(field_name) = property_field_name(body, field_names) else {
        return body.to_vec();
    };
    vec![Statement::new(StmtKind::Return(Some(Expression::new(ExprKind::Member {
        object: Box::new(Expression::new(ExprKind::This)),
        field: field_name,
        null_safe: false,
    }))))]
}

fn rewrite_property_setter_body(body: &[Statement], field_names: &HashSet<String>) -> Vec<Statement> {
    let Some(field_name) = property_field_name(body, field_names) else {
        return body.to_vec();
    };
    vec![Statement::new(StmtKind::Assign {
        targets: vec![Expression::new(ExprKind::Member {
            object: Box::new(Expression::new(ExprKind::This)),
            field: field_name,
            null_safe: false,
        })],
        value: Expression::ident("value"),
    })]
}

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
    let mut constructors: Vec<NormalConstructor> = Vec::new();
    let mut destructor: Option<NormalMethod> = None;
    let mut special_methods: Vec<SpecialMethod> = Vec::new();
    let field_names: HashSet<String> = members.iter().filter_map(|member| match member {
        ClassMember::Field { name, .. } => Some(name.to_ascii_lowercase()),
        _ => None,
    }).collect();

    for member in members {
        match member {
            ClassMember::Field { name: fname, type_hint, init, modifiers: m, .. } => {
                let field = NormalField {
                    span: span.clone(),
                    name: fname.clone(),
                    type_hint: type_hint.clone(),
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
                    if let Some(d) = from_method_stmt(span.clone(), stmt, "destroy", Access::Public) {
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
                constructors.push(NormalConstructor {
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
            ClassMember::Property { name: pname, getter, setter, is_auto, modifiers: m, .. } => {
                let (canonical, _) = canonicalize_method(ClassLang::Pascal, pname);
                let getter_method = getter.as_ref().map(|body| build_normal_method(
                    span.clone(), &canonical, pname, Vec::new(),
                    vec![], None, rewrite_property_getter_body(body, &field_names),
                    Access::Public, false, false, false, Modifiers::default(),
                ));
                let setter_method = setter.as_ref().map(|s: &PropertySetter| build_normal_method(
                    span.clone(), &canonical, pname, Vec::new(),
                    vec![s.param.clone()], None, rewrite_property_setter_body(&s.body, &field_names),
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

    instance_methods = lower_pascal_method_overloads(instance_methods, &span);
    static_methods = lower_pascal_method_overloads(static_methods, &span);
    let (ctor_helper_methods, constructor) = lower_pascal_constructor_overloads(constructors, &span);
    instance_methods.extend(ctor_helper_methods);

    if let Some(destructor_method) = destructor.clone() {
        instance_methods.push(destructor_method);

        let has_free = instance_methods.iter().any(|method| {
            method.source_name.eq_ignore_ascii_case("Free")
                || method.canonical_name.eq_ignore_ascii_case("free")
        });
        if !has_free {
            let free_body = vec![Statement::new(StmtKind::Expr(Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident("Destroy")),
                args: Vec::new(),
                optional: false,
            })))];
            instance_methods.push(build_normal_method(
                span.clone(),
                "free",
                "Free",
                Vec::new(),
                Vec::new(),
                None,
                free_body,
                Access::Public,
                false,
                false,
                true,
                Modifiers::default(),
            ));
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
        explicit_self_param: false, // Pascal: Self is implicit
        implicit_self_fields: true, // Pascal: bare field names resolve to Self.field inside methods
        instance_fields,
        static_fields,
        instance_methods,
        static_methods,
        properties,
        constructors: Vec::new(),
        constructor,
        destructor,
        auto_init_methods: Vec::new(),
        special_methods,
        event_bindings: Vec::new(),
        raw_extra_members,
    }
}

#[derive(Clone)]
struct PascalOverloadCase {
    arity: usize,
    hidden_name: String,
}

fn lower_pascal_method_overloads(methods: Vec<NormalMethod>, span: &Span) -> Vec<NormalMethod> {
    let mut groups: HashMap<String, Vec<NormalMethod>> = HashMap::new();
    let mut order = Vec::new();

    for method in methods {
        if !groups.contains_key(&method.canonical_name) {
            order.push(method.canonical_name.clone());
        }
        groups.entry(method.canonical_name.clone()).or_default().push(method);
    }

    let mut lowered = Vec::new();
    for key in order {
        let Some(group) = groups.remove(&key) else { continue; };
        if group.len() <= 1 || has_duplicate_arities(group.iter().map(|m| m.params.len())) {
            lowered.extend(group);
            continue;
        }

        let mut hidden_methods = Vec::new();
        let mut cases = Vec::new();
        let mut sorted = group;
        sorted.sort_by_key(|method| method.params.len());
        let wrapper_template = sorted.last().cloned().unwrap();

        for method in sorted {
            let hidden_name = format!("__vybe_overload_{}_{}", method.canonical_name, method.params.len());
            cases.push(PascalOverloadCase {
                arity: method.params.len(),
                hidden_name: hidden_name.clone(),
            });
            hidden_methods.push(build_normal_method(
                method.span.clone(),
                &hidden_name,
                &hidden_name,
                Vec::new(),
                method.params.clone(),
                method.return_type.clone(),
                method.body.clone(),
                method.access,
                method.is_async,
                method.is_generator,
                method.is_sub,
                method.raw_modifiers.clone(),
            ));
        }

        lowered.extend(hidden_methods);
        lowered.push(build_normal_method(
            span.clone(),
            &wrapper_template.canonical_name,
            &wrapper_template.source_name,
            wrapper_template.aliases.clone(),
            wrapper_template.params.clone(),
            wrapper_template.return_type.clone(),
            build_pascal_overload_dispatch(&cases, &wrapper_template.params, wrapper_template.return_type.is_none() && wrapper_template.is_sub),
            wrapper_template.access,
            false,
            false,
            wrapper_template.is_sub,
            Modifiers::default(),
        ));
    }

    lowered
}

fn lower_pascal_constructor_overloads(
    constructors: Vec<NormalConstructor>,
    span: &Span,
) -> (Vec<NormalMethod>, Option<NormalConstructor>) {
    if constructors.is_empty() {
        return (Vec::new(), None);
    }
    if constructors.len() == 1 || has_duplicate_arities(constructors.iter().map(|ctor| ctor.params.len())) {
        return (Vec::new(), constructors.into_iter().last());
    }

    let mut sorted = constructors;
    sorted.sort_by_key(|ctor| ctor.params.len());
    let wrapper_template = sorted.last().cloned().unwrap();
    let mut helper_methods = Vec::new();
    let mut cases = Vec::new();

    for ctor in sorted {
        let hidden_name = format!("__vybe_ctor_create_{}", ctor.params.len());
        cases.push(PascalOverloadCase {
            arity: ctor.params.len(),
            hidden_name: hidden_name.clone(),
        });
        helper_methods.push(build_normal_method(
            ctor.span.clone(),
            &hidden_name,
            &hidden_name,
            Vec::new(),
            ctor.params.clone(),
            None,
            ctor.body.clone(),
            Access::Public,
            false,
            false,
            true,
            Modifiers::default(),
        ));
    }

    let wrapper = NormalConstructor {
        span: span.clone(),
        params: wrapper_template.params.clone(),
        body: build_pascal_overload_dispatch(&cases, &wrapper_template.params, true),
        base_call: wrapper_template.base_call,
        named_name: wrapper_template.named_name,
    };

    (helper_methods, Some(wrapper))
}

fn build_pascal_overload_dispatch(
    cases: &[PascalOverloadCase],
    wrapper_params: &[crate::ast::Param],
    is_sub: bool,
) -> Vec<Statement> {
    if cases.is_empty() {
        return Vec::new();
    }

    let first = &cases[0];
    let call_args: Vec<crate::ast::Argument> = wrapper_params.iter().take(first.arity)
        .map(|param| crate::ast::Argument::positional(Expression::ident(&param.name)))
        .collect();
    let call_expr = Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident(&first.hidden_name)),
        args: call_args,
        optional: false,
    });
    let invoke_stmt = if is_sub {
        Statement::new(StmtKind::Expr(call_expr))
    } else {
        Statement::new(StmtKind::Return(Some(call_expr)))
    };

    if cases.len() == 1 {
        return vec![invoke_stmt];
    }

    let gate_param = &wrapper_params[first.arity].name;
    let cond = Expression::new(ExprKind::Binary {
        op: crate::ast::BinOp::Eq,
        left: Box::new(Expression::ident(gate_param)),
        right: Box::new(Expression::null()),
    });

    vec![Statement::new(StmtKind::If {
        cond,
        then_body: vec![invoke_stmt],
        elifs: Vec::new(),
        else_body: Some(build_pascal_overload_dispatch(&cases[1..], wrapper_params, is_sub)),
    })]
}

fn has_duplicate_arities<I>(arities: I) -> bool
where
    I: IntoIterator<Item = usize>,
{
    let mut seen = std::collections::HashSet::new();
    for arity in arities {
        if !seen.insert(arity) {
            return true;
        }
    }
    false
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::ast::Modifiers;

    fn dummy_span() -> Span { Span::default() }

    fn make_method(src_name: &str) -> ClassMember {
        ClassMember::Method(Box::new(crate::ast::Statement::new(StmtKind::FunctionDecl {
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

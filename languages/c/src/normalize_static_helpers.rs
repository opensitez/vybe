//! Give parameter-only generated IIFEs a statically resolved callable target.

use std::collections::{HashMap, HashSet};
use vybe_ast::{ExprKind, Expression, LambdaBody, ObjectProperty, Statement, StmtKind};

#[path = "static_helper_key.rs"]
mod helper_key;

#[cfg(test)]
pub fn normalize(body: &mut Vec<Statement>, mut helpers: HashSet<String>) {
    normalize_named(body, &mut helpers, HashSet::new());
}

pub fn normalize_with_imports(
    body: &mut Vec<Statement>,
    mut helpers: HashSet<String>,
    imports: &[vybe_ast::Import],
) {
    use vybe_ast::ImportKind;
    let mut names = HashSet::new();
    for import in imports {
        match &import.kind {
            ImportKind::Named {
                names: imported, ..
            } => {
                names.extend(
                    imported
                        .iter()
                        .map(|name| name.alias.as_ref().unwrap_or(&name.name).clone()),
                );
            }
            ImportKind::Default { local, .. } => {
                names.insert(local.clone());
            }
            ImportKind::Simple { path, alias } | ImportKind::Wildcard { path, alias } => {
                names.insert(
                    alias
                        .clone()
                        .unwrap_or_else(|| path.rsplit('.').next().unwrap_or(path).into()),
                );
            }
        }
    }
    normalize_named(body, &mut helpers, names.clone());
    normalize_fixed_with_reserved(body, &helpers, names);
}

fn normalize_named(
    body: &mut Vec<Statement>,
    helpers: &mut HashSet<String>,
    mut declared_names: HashSet<String>,
) {
    let mut declarations: Vec<Statement> = Vec::new();
    let mut implementations: HashMap<u64, Vec<usize>> = HashMap::new();
    reserve_bindings(body, &mut declared_names);
    let mut next_name = 0;
    for statement in body.iter_mut() {
        statement.walk_exprs_mut(&mut |value| {
            let ExprKind::Call {
                callee,
                args,
                optional: false,
            } = &mut value.kind
            else {
                return;
            };
            let ExprKind::Lambda {
                params,
                body: LambdaBody::Expr(result),
                is_async: false,
                captures,
            } = &callee.kind
            else {
                return;
            };
            if !captures.is_empty()
                || args.len() != params.len()
                || args
                    .iter()
                    .any(|arg| arg.spread || arg.by_ref || arg.name.is_some())
                || params.iter().any(|param| {
                    param.type_hint.is_some()
                        || param.default.is_some()
                        || param.is_rest
                        || param.is_kwargs
                        || param.is_optional
                        || param.pass_by != vybe_ast::PassBy::Value
                })
            {
                return;
            }
            let raw_arguments = args.iter().any(|arg| !primitive_argument(&arg.value));
            if raw_arguments && (args.len() > 254 || declared_names.contains("__c_invoke_fixed")) {
                return;
            }
            let names = params.iter().map(|param| param.name.clone()).collect();
            if !parameter_only(result, &names, helpers) {
                return;
            }
            // Only these newly generated, non-escaping immediate-call targets
            // may share an implementation. Never merge source declarations or
            // exposed function values. Hash collisions require a full match.
            let fingerprint = helper_key::fingerprint(params, result);
            let existing = fingerprint
                .and_then(|key| implementations.get(&key))
                .and_then(|indices| indices.iter().find_map(|&index| {
                    let StmtKind::FunctionDecl { name, params: other_params, body, .. } =
                        &declarations[index].kind else { return None; };
                    let [statement] = body.as_slice() else { return None; };
                    let StmtKind::Return(Some(other)) = &statement.kind else { return None; };
                    helper_key::equivalent(params, result, other_params, other)
                        .then(|| name.clone())
                }));
            if let Some(name) = existing {
                replace_generated_target(callee, args, name, raw_arguments);
                return;
            }
            let name = loop {
                let name = format!("__c_static_iife_{next_name}");
                next_name += 1;
                if declared_names.insert(name.clone()) {
                    break name;
                }
            };
            let ExprKind::Lambda { params, body: LambdaBody::Expr(result), .. } =
                std::mem::replace(&mut callee.kind, ExprKind::Lit(vybe_ast::Literal::Int(0)))
            else { unreachable!("eligibility already checked the literal lambda") };
            let mut declaration = vybe_platform_libc::emitter::build::function_stmt(
                &name,
                Vec::new(),
                vec![Statement::new(StmtKind::Return(Some(*result)))],
            );
            if let StmtKind::FunctionDecl {
                params: declared, ..
            } = &mut declaration.kind
            {
                *declared = params;
            }
            if let Some(key) = fingerprint {
                implementations.entry(key).or_default().push(declarations.len());
            }
            declarations.push(declaration);
            helpers.insert(name.clone());
            replace_generated_target(callee, args, name, raw_arguments);
        });
    }
    declarations.append(body);
    *body = declarations;
}

fn replace_generated_target(
    callee: &mut Expression,
    args: &mut Vec<vybe_ast::Argument>,
    name: String,
    raw_arguments: bool,
) {
    if raw_arguments {
        // A normal source call may copy a managed C record by value. These
        // generated parameters instead retain the literal lambda's raw-value
        // contract through the existing common callable invocation emitter.
        args.insert(0, vybe_ast::Argument::positional(Expression {
            kind: ExprKind::Ident(name),
            span: callee.span,
        }));
        callee.kind = ExprKind::Ident("__c_invoke_fixed".into());
    } else {
        callee.kind = ExprKind::Ident(name);
    }
}

/// Keep the literal callable and its raw-value arguments, but omit generic
/// callable classification, rest-arity and dynamic receiver dispatch.
#[cfg(test)]
pub fn normalize_fixed_invocations(body: &mut [Statement], helpers: &HashSet<String>) {
    normalize_fixed_with_reserved(body, helpers, HashSet::new());
}

fn normalize_fixed_with_reserved(
    body: &mut [Statement],
    helpers: &HashSet<String>,
    mut names: HashSet<String>,
) {
    reserve_bindings(body, &mut names);
    if names.contains("__c_invoke_fixed") {
        return;
    }
    for statement in body {
        statement.walk_exprs_mut(&mut |value| {
            let ExprKind::Call {
                callee,
                args,
                optional: false,
            } = &mut value.kind
            else {
                return;
            };
            let ExprKind::Lambda {
                params,
                body: LambdaBody::Expr(result),
                is_async: false,
                captures,
            } = &callee.kind
            else {
                return;
            };
            if !captures.is_empty()
                || args.len() != params.len()
                // Leave room for both the callable builtin argument and the
                // module ABI's implicit receiver in the u8 arity contract.
                || args.len() > 254
                || args.iter().any(|arg| arg.spread || arg.by_ref || arg.name.is_some())
                || params.iter().any(|param| {
                    param.type_hint.is_some() || param.default.is_some()
                        || param.is_rest || param.is_kwargs || param.is_optional
                        || param.pass_by != vybe_ast::PassBy::Value
                })
            {
                return;
            }
            let params = params.iter().map(|param| param.name.clone()).collect();
            if !parameter_only(result, &params, helpers) {
                return;
            }
            let span = callee.span;
            let literal = std::mem::replace(
                callee,
                Box::new(Expression {
                    kind: ExprKind::Ident("__c_invoke_fixed".into()),
                    span,
                }),
            );
            args.insert(0, vybe_ast::Argument::positional(*literal));
        });
    }
}

fn reserve_pattern(pattern: &vybe_ast::BindingPattern, names: &mut HashSet<String>) {
    use vybe_ast::{ArrayPatternElem, BindingPattern};
    match pattern {
        BindingPattern::Ident(name) => {
            names.insert(name.clone());
        }
        BindingPattern::Object(properties) => {
            for property in properties {
                if let Some(value) = &property.value {
                    reserve_pattern(value, names);
                } else {
                    names.insert(property.key.clone());
                }
            }
        }
        BindingPattern::Array(elements) => {
            for element in elements {
                match element {
                    ArrayPatternElem::Pattern(pattern, _) => reserve_pattern(pattern, names),
                    ArrayPatternElem::Rest(name) => {
                        names.insert(name.clone());
                    }
                    ArrayPatternElem::Hole => {}
                }
            }
        }
    }
}

// Reserve local names too: a newly introduced global helper reference must
// not resolve to an existing local, even when that local was never read.
fn reserve_bindings(body: &[Statement], names: &mut HashSet<String>) {
    for statement in body {
        statement.walk_exprs(&mut |expression| match &expression.kind {
            ExprKind::Ident(name) => {
                names.insert(name.clone());
            }
            ExprKind::Lambda {
                params,
                captures,
                body,
                ..
            } => {
                names.extend(params.iter().map(|param| param.name.clone()));
                names.extend(captures.iter().cloned());
                if let LambdaBody::Block(body) = body {
                    reserve_bindings(body, names);
                }
            }
            ExprKind::FunctionExpr(declaration) => {
                reserve_bindings(std::slice::from_ref(declaration.as_ref()), names);
            }
            _ => {}
        });
        match &statement.kind {
            StmtKind::VarDecl { declarations, .. } => {
                for declaration in declarations {
                    reserve_pattern(&declaration.pattern, names);
                }
            }
            StmtKind::FunctionDecl {
                name, params, body, ..
            } => {
                names.insert(name.clone());
                names.extend(params.iter().map(|param| param.name.clone()));
                reserve_bindings(body, names);
            }
            StmtKind::ClassDecl { name, members, .. }
            | StmtKind::StructDecl { name, members, .. }
            | StmtKind::ModuleDecl { name, members, .. } => {
                names.insert(name.clone());
                for member in members {
                    match member {
                        vybe_ast::ClassMember::Method(declaration)
                        | vybe_ast::ClassMember::NestedType(declaration) => {
                            reserve_bindings(std::slice::from_ref(declaration.as_ref()), names);
                        }
                        vybe_ast::ClassMember::Constructor { params, body, .. } => {
                            names.extend(params.iter().map(|param| param.name.clone()));
                            reserve_bindings(body, names);
                        }
                        vybe_ast::ClassMember::Property { getter, setter, .. } => {
                            if let Some(body) = getter {
                                reserve_bindings(body, names);
                            }
                            if let Some(setter) = setter {
                                names.insert(setter.param.name.clone());
                                reserve_bindings(&setter.body, names);
                            }
                        }
                        _ => {}
                    }
                }
            }
            StmtKind::InterfaceDecl { name, .. } | StmtKind::DelegateDecl { name, .. } => {
                names.insert(name.clone());
            }
            StmtKind::EnumDecl { name, members, .. } => {
                names.insert(name.clone());
                names.extend(members.iter().map(|member| member.name.clone()));
            }
            StmtKind::NamespaceDecl { name, body } => {
                names.insert(name.clone());
                reserve_bindings(body, names);
            }
            StmtKind::Block(body)
            | StmtKind::DoWhile { body, .. }
            | StmtKind::Lock { body, .. } => reserve_bindings(body, names),
            StmtKind::If {
                then_body,
                elifs,
                else_body,
                ..
            } => {
                reserve_bindings(then_body, names);
                for (_, body) in elifs {
                    reserve_bindings(body, names);
                }
                if let Some(body) = else_body {
                    reserve_bindings(body, names);
                }
            }
            StmtKind::For { init, body, .. } => {
                if let Some(init) = init {
                    reserve_bindings(std::slice::from_ref(init.as_ref()), names);
                }
                reserve_bindings(body, names);
            }
            StmtKind::ForIn {
                var,
                key,
                body,
                else_body,
                ..
            } => {
                names.insert(var.clone());
                names.extend(key.iter().cloned());
                reserve_bindings(body, names);
                if let Some(body) = else_body {
                    reserve_bindings(body, names);
                }
            }
            StmtKind::While {
                body, else_body, ..
            } => {
                reserve_bindings(body, names);
                if let Some(body) = else_body {
                    reserve_bindings(body, names);
                }
            }
            StmtKind::Switch { cases, default, .. } => {
                for case in cases {
                    reserve_bindings(&case.body, names);
                }
                if let Some(body) = default {
                    reserve_bindings(body, names);
                }
            }
            StmtKind::Try {
                body,
                catches,
                else_body,
                finally,
            } => {
                reserve_bindings(body, names);
                for catch in catches {
                    names.extend(catch.var_name.iter().cloned());
                    names.extend(catch.stack_var.iter().cloned());
                    reserve_bindings(&catch.body, names);
                }
                for body in else_body.iter().chain(finally.iter()) {
                    reserve_bindings(body, names);
                }
            }
            StmtKind::With { items, body, .. } => {
                names.extend(items.iter().filter_map(|item| item.var.clone()));
                reserve_bindings(body, names);
            }
            StmtKind::Using { var, body, .. } => {
                names.insert(var.clone());
                reserve_bindings(body, names);
            }
            StmtKind::ScopeDecl {
                names: declared, ..
            } => {
                names.extend(declared.iter().cloned());
            }
            StmtKind::Export {
                declaration: Some(body),
                ..
            }
            | StmtKind::Labeled { body, .. } => {
                reserve_bindings(std::slice::from_ref(body.as_ref()), names);
            }
            _ => {}
        }
    }
}

// Do not infer backing/copy safety from an expression's outer shape. Only
// primitive literals, predicates and explicit integer-lane operations prove
// a non-record result; casts and arbitrary call results remain unproven.
fn primitive_argument(value: &Expression) -> bool {
    use vybe_ast::{BinOp, Literal, UnaryOp};
    match &value.kind {
        ExprKind::Lit(
            Literal::Int(_)
            | Literal::Float(_)
            | Literal::Bool(_)
            | Literal::Char(_)
            | Literal::Str(_)
            | Literal::Null
            | Literal::Undefined,
        ) => true,
        ExprKind::IsType { .. } => true,
        ExprKind::Unary {
            op: UnaryOp::Not | UnaryOp::Typeof,
            ..
        } => true,
        ExprKind::Binary {
            op: BinOp::Integer(_, _),
            ..
        } => true,
        ExprKind::Sequence(values) => values.last().is_some_and(primitive_argument),
        ExprKind::Ternary { then, else_, .. } => {
            primitive_argument(then) && primitive_argument(else_)
        }
        _ => false,
    }
}

// A whitelist is intentional: an unrecognised node must not hide a storage
// reference, a closure, or an implicit receiver from this capture proof.
fn parameter_only(value: &Expression, params: &HashSet<String>, helpers: &HashSet<String>) -> bool {
    let child = |value: &Expression| parameter_only(value, params, helpers);
    match &value.kind {
        ExprKind::Lit(_) => true,
        ExprKind::Ident(name) => params.contains(name),
        ExprKind::Member { object, .. } => child(object),
        ExprKind::Index { object, index, .. } => child(object) && child(index),
        ExprKind::Unary { expr, op }
            if matches!(
                op,
                vybe_ast::UnaryOp::Not | vybe_ast::UnaryOp::Neg | vybe_ast::UnaryOp::Typeof
            ) =>
        {
            child(expr)
        }
        ExprKind::Cast { expr, .. } | ExprKind::IsType { expr, .. } => child(expr),
        ExprKind::Binary { left, right, .. } => child(left) && child(right),
        ExprKind::Ternary { cond, then, else_ } => child(cond) && child(then) && child(else_),
        ExprKind::Sequence(values) => values.iter().all(child),
        ExprKind::Object(properties) => properties.iter().all(|property| match property {
            ObjectProperty::KeyValue { key, value } | ObjectProperty::Computed { key, value } => {
                child(key) && child(value)
            }
            _ => false,
        }),
        ExprKind::Call {
            callee,
            args,
            optional: false,
        } => {
            matches!(&callee.kind, ExprKind::Ident(name) if helpers.contains(name) && !params.contains(name))
                && args.iter().all(|arg| {
                    !arg.spread && !arg.by_ref && arg.name.is_none() && child(&arg.value)
                })
        }
        _ => false,
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use vybe_ast::{Argument, BinOp, BitLane, IntOp, Literal};
    use vybe_platform_libc::emitter::build::function_stmt;

    fn invocation(result: Expression) -> Statement {
        let StmtKind::FunctionDecl { params, .. } =
            function_stmt("unused", vec!["operand"], vec![]).kind
        else {
            unreachable!()
        };
        Statement::new(StmtKind::Expr(Expression::new(ExprKind::Call {
            callee: Box::new(Expression::new(ExprKind::Lambda {
                params,
                body: LambdaBody::Expr(Box::new(result)),
                is_async: false,
                captures: vec![],
            })),
            args: vec![Argument::positional(Expression::new(ExprKind::Binary {
                op: BinOp::Integer(IntOp::Add, BitLane::W32),
                left: Box::new(Expression::new(ExprKind::Call {
                    callee: Box::new(Expression::ident("advance")),
                    args: vec![],
                    optional: false,
                })),
                right: Box::new(Expression::new(ExprKind::Lit(Literal::Int(0)))),
            }))],
            optional: false,
        })))
    }

    #[test]
    fn lifts_parameter_only_body_without_repeating_argument() {
        let mut body = vec![invocation(Expression::new(ExprKind::Binary {
            op: BinOp::Eq,
            left: Box::new(Expression::ident("operand")),
            right: Box::new(Expression::ident("operand")),
        }))];
        normalize(&mut body, HashSet::new());
        assert_eq!(body.len(), 2);
        let StmtKind::FunctionDecl {
            params,
            body: helper_body,
            ..
        } = &body[0].kind
        else {
            panic!()
        };
        assert_eq!(params[0].name, "operand");
        assert!(matches!(helper_body[0].kind, StmtKind::Return(Some(_))));
        let StmtKind::Expr(value) = &body[1].kind else {
            panic!()
        };
        let ExprKind::Call { callee, args, .. } = &value.kind else {
            panic!()
        };
        assert!(matches!(&callee.kind, ExprKind::Ident(name) if name == "__c_static_iife_0"));
        assert_eq!(args.len(), 1);
        assert!(matches!(&args[0].value.kind, ExprKind::Binary { left, .. }
            if matches!(&left.kind, ExprKind::Call { callee, .. }
                if matches!(&callee.kind, ExprKind::Ident(name) if name == "advance"))));
    }

    #[test]
    fn shares_only_generated_immediate_targets_and_keeps_both_calls() {
        let mut first = invocation(Expression::ident("operand"));
        let mut second = first.clone();
        first.walk_exprs_mut(&mut |value| value.span = Default::default());
        second.walk_exprs_mut(&mut |value| value.span.start_line = 123);
        let mut body = vec![first, second, invocation(Expression::int(7))];
        normalize(&mut body, HashSet::new());
        assert_eq!(body.len(), 5, "two distinct implementations and three calls");
        let targets: Vec<_> = body[2..].iter().map(|statement| {
            let StmtKind::Expr(value) = &statement.kind else { panic!() };
            let ExprKind::Call { callee, args, .. } = &value.kind else { panic!() };
            assert_eq!(args.len(), 1);
            assert!(matches!(&args[0].value.kind, ExprKind::Binary { left, .. }
                if matches!(&left.kind, ExprKind::Call { .. })));
            let ExprKind::Ident(name) = &callee.kind else { panic!() };
            name.clone()
        }).collect();
        assert_eq!(targets[0], targets[1]);
        assert_ne!(targets[0], targets[2]);
    }

    #[test]
    fn preserves_existing_helper_names_and_is_idempotent() {
        let mut body = vec![
            function_stmt("__c_static_iife_0", vec![], vec![]),
            invocation(Expression::ident("operand")),
        ];
        normalize(&mut body, HashSet::new());
        assert_eq!(body.len(), 3);
        assert!(matches!(&body[0].kind, StmtKind::FunctionDecl { name, .. }
            if name == "__c_static_iife_1"));
        normalize(&mut body, HashSet::new());
        assert_eq!(body.len(), 3);
    }

    #[test]
    fn reserves_global_variable_and_nested_pattern_bindings() {
        use vybe_ast::{
            ArrayPatternElem, BindingPattern, ObjectPatternProp, VarDeclKind, VarDeclarator,
        };
        let declaration = Statement::new(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Array(vec![
                    ArrayPatternElem::Pattern(
                        BindingPattern::Ident("__c_static_iife_0".into()),
                        None,
                    ),
                    ArrayPatternElem::Pattern(
                        BindingPattern::Object(vec![ObjectPatternProp {
                            key: "__c_static_iife_1".into(),
                            value: None,
                            default: None,
                            is_rest: false,
                        }]),
                        None,
                    ),
                    ArrayPatternElem::Rest("__c_static_iife_2".into()),
                ]),
                type_hint: None,
                init: None,
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        });
        let original = format!("{declaration:?}");
        let mut body = vec![declaration, invocation(Expression::ident("operand"))];
        normalize(&mut body, HashSet::new());
        assert!(matches!(&body[0].kind, StmtKind::FunctionDecl { name, .. }
            if name == "__c_static_iife_3"));
        assert_eq!(format!("{:?}", body[1]), original);
    }

    #[test]
    fn reserves_unused_global_and_local_variable_names() {
        use vybe_ast::{BindingPattern, VarDeclKind, VarDeclarator};
        let variable = |name: &str| {
            Statement::new(StmtKind::VarDecl {
                declarations: vec![VarDeclarator {
                    pattern: BindingPattern::Ident(name.into()),
                    type_hint: None,
                    init: Some(Expression::new(ExprKind::Lit(Literal::Int(7)))),
                    array_bounds: None,
                    with_events: false,
                }],
                kind: VarDeclKind::Var,
            })
        };
        let mut body = vec![
            variable("__c_static_iife_0"),
            function_stmt(
                "caller",
                vec!["__c_static_iife_1"],
                vec![
                    variable("__c_static_iife_2"),
                    invocation(Expression::ident("operand")),
                ],
            ),
        ];
        let original_global = format!("{:?}", body[0]);
        normalize(&mut body, HashSet::new());
        assert!(matches!(&body[0].kind, StmtKind::FunctionDecl { name, .. }
            if name == "__c_static_iife_3"));
        assert_eq!(format!("{:?}", body[1]), original_global);
        let StmtKind::FunctionDecl {
            params,
            body: caller_body,
            ..
        } = &body[2].kind
        else {
            panic!()
        };
        assert_eq!(params[0].name, "__c_static_iife_1");
        assert!(
            matches!(&caller_body[0].kind, StmtKind::VarDecl { declarations, .. }
            if matches!(&declarations[0].pattern, BindingPattern::Ident(name)
                if name == "__c_static_iife_2"))
        );
    }

    #[test]
    fn reserves_imported_binding_aliases() {
        let mut body = vec![invocation(Expression::ident("operand"))];
        let imports = vec![vybe_ast::Import {
            kind: vybe_ast::ImportKind::Named {
                path: "support".into(),
                names: vec![vybe_ast::ImportName {
                    name: "external".into(),
                    alias: Some("__c_static_iife_0".into()),
                }],
                level: 0,
            },
            span: Default::default(),
        }];
        normalize_with_imports(&mut body, HashSet::new(), &imports);
        assert!(matches!(&body[0].kind, StmtKind::FunctionDecl { name, .. }
            if name == "__c_static_iife_1"));
    }

    #[test]
    fn fixed_invocation_retains_computed_managed_arguments_without_copy() {
        for argument in [
            Expression::ident("record"),
            Expression::new(ExprKind::Cast {
                expr: Box::new(Expression::ident("record")),
                type_name: "node".into(),
            }),
            Expression::new(ExprKind::Sequence(vec![
                Expression::new(ExprKind::Lit(Literal::Int(0))),
                Expression::ident("record"),
            ])),
            Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident("managed_result")),
                args: vec![],
                optional: false,
            }),
        ] {
            let mut statement = invocation(Expression::ident("operand"));
            let StmtKind::Expr(value) = &mut statement.kind else {
                panic!()
            };
            let ExprKind::Call { args, .. } = &mut value.kind else {
                panic!()
            };
            args[0].value = argument;
            let original_argument = format!("{:?}", args[0]);
            let mut body = vec![statement];
            normalize_fixed_invocations(&mut body, &HashSet::new());
            assert_eq!(body.len(), 1);
            let StmtKind::Expr(value) = &body[0].kind else {
                panic!()
            };
            let ExprKind::Call { callee, args, .. } = &value.kind else {
                panic!()
            };
            assert!(matches!(&callee.kind, ExprKind::Ident(name) if name == "__c_invoke_fixed"));
            assert!(matches!(&args[0].value.kind, ExprKind::Lambda { .. }));
            assert_eq!(args.len(), 2);
            assert_eq!(format!("{:?}", args[1]), original_argument);
        }
    }

    #[test]
    fn fixed_invocation_rejects_implicit_receiver_and_builtin_shadowing() {
        let mut body = vec![invocation(Expression::new(ExprKind::This))];
        let original = format!("{body:?}");
        normalize_fixed_invocations(&mut body, &HashSet::new());
        assert_eq!(format!("{body:?}"), original);

        let mut body = vec![function_stmt(
            "caller",
            vec!["__c_invoke_fixed"],
            vec![invocation(Expression::ident("operand"))],
        )];
        let original = format!("{body:?}");
        normalize_fixed_invocations(&mut body, &HashSet::new());
        assert_eq!(format!("{body:?}"), original);
    }

    #[test]
    fn leaves_free_variables_and_storage_references_in_scope() {
        for result in [
            Expression::ident("outer"),
            Expression::new(ExprKind::RefLoad(Box::new(Expression::ident("operand")))),
        ] {
            let mut body = vec![invocation(result)];
            normalize(&mut body, HashSet::new());
            assert_eq!(body.len(), 1);
        }
    }

    #[test]
    fn leaves_potential_record_arguments_on_the_no_copy_path() {
        for argument in [
            Expression::ident("record"),
            Expression::new(ExprKind::Member {
                object: Box::new(Expression::ident("record")),
                field: "backing".into(),
                null_safe: false,
            }),
            Expression::new(ExprKind::Index {
                object: Box::new(Expression::ident("records")),
                index: Box::new(Expression::new(ExprKind::Lit(Literal::Int(0)))),
                null_safe: false,
            }),
            Expression::new(ExprKind::Cast {
                expr: Box::new(Expression::ident("record")),
                type_name: "record".into(),
            }),
            Expression::new(ExprKind::Sequence(vec![
                Expression::new(ExprKind::Lit(Literal::Int(0))),
                Expression::ident("record"),
            ])),
            Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident("managed_result")),
                args: vec![],
                optional: false,
            }),
        ] {
            let mut statement = invocation(Expression::ident("operand"));
            let StmtKind::Expr(value) = &mut statement.kind else {
                panic!()
            };
            let ExprKind::Call { args, .. } = &mut value.kind else {
                panic!()
            };
            args[0].value = argument;
            let original_argument = format!("{:?}", args[0]);
            let mut body = vec![statement];
            normalize(&mut body, HashSet::new());
            assert_eq!(body.len(), 2);
            let StmtKind::FunctionDecl { name, .. } = &body[0].kind else { panic!() };
            let StmtKind::Expr(value) = &body[1].kind else { panic!() };
            let ExprKind::Call { callee, args, .. } = &value.kind else { panic!() };
            assert!(matches!(&callee.kind, ExprKind::Ident(name) if name == "__c_invoke_fixed"));
            assert_eq!(args.len(), 2);
            assert!(matches!(&args[0].value.kind, ExprKind::Ident(target) if target == name));
            assert_eq!(format!("{:?}", args[1]), original_argument);
        }
    }

    #[test]
    fn managed_target_sharing_retains_calls_and_respects_reserved_builtin() {
        let make = || {
            let mut statement = invocation(Expression::ident("operand"));
            let StmtKind::Expr(value) = &mut statement.kind else { panic!() };
            let ExprKind::Call { args, .. } = &mut value.kind else { panic!() };
            args[0].value = Expression::ident("record");
            statement
        };
        let mut body = vec![make(), make()];
        normalize(&mut body, HashSet::new());
        assert_eq!(body.len(), 3);
        let targets: Vec<_> = body[1..].iter().map(|statement| {
            let StmtKind::Expr(value) = &statement.kind else { panic!() };
            let ExprKind::Call { callee, args, .. } = &value.kind else { panic!() };
            assert!(matches!(&callee.kind, ExprKind::Ident(name) if name == "__c_invoke_fixed"));
            assert_eq!(args.len(), 2);
            let ExprKind::Ident(name) = &args[0].value.kind else { panic!() };
            assert!(matches!(&args[1].value.kind, ExprKind::Ident(name) if name == "record"));
            name.clone()
        }).collect();
        assert_eq!(targets[0], targets[1]);
        let original = format!("{body:?}");
        normalize(&mut body, HashSet::new());
        assert_eq!(format!("{body:?}"), original);

        let mut body = vec![function_stmt("caller", vec!["__c_invoke_fixed"], vec![make()])];
        let original = format!("{body:?}");
        normalize(&mut body, HashSet::new());
        assert_eq!(format!("{body:?}"), original);
    }

    #[test]
    fn only_reserved_support_calls_can_be_free_identifiers() {
        for (name, expected) in [("__c_ref_eq", 2), ("outer_callable", 1)] {
            let mut body = vec![invocation(Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident(name)),
                args: vec![Argument::positional(Expression::ident("operand"))],
                optional: false,
            }))];
            normalize(&mut body, HashSet::from(["__c_ref_eq".into()]));
            assert_eq!(body.len(), expected);
        }
    }

    #[test]
    fn keeps_spread_and_optional_calls_dynamic() {
        for spread in [false, true] {
            let mut statement = invocation(Expression::new(ExprKind::Lit(Literal::Int(1))));
            let StmtKind::Expr(value) = &mut statement.kind else {
                panic!()
            };
            let ExprKind::Call { args, optional, .. } = &mut value.kind else {
                panic!()
            };
            args[0].spread = spread;
            *optional = !spread;
            let mut body = vec![statement];
            normalize(&mut body, HashSet::new());
            assert_eq!(body.len(), 1);
        }
    }

    #[test]
    fn static_target_omits_generic_invocation_opcodes() {
        crate::register();
        let mut module =
            crate::parse("int calls = 0; int advance(void) { ++calls; return 1; } int main(void) { return calls != 2; }").unwrap();
        let main = module.body.iter_mut().find(|statement| {
            matches!(&statement.kind, StmtKind::FunctionDecl { name, .. } if name == "main")
        }).unwrap();
        let StmtKind::FunctionDecl { body, .. } = &mut main.kind else {
            unreachable!()
        };
        body.insert(0, invocation(Expression::ident("operand")));
        body.insert(0, invocation(Expression::ident("operand")));
        let profile = vybe_runtime::profile::parse_profile_shared(crate::profile_source()).unwrap();
        let baseline = vybe_compiler::primitives::Compiler::with_shared_profile(profile.clone())
            .compile(&module)
            .unwrap();
        normalize(&mut module.body, HashSet::new());
        assert_eq!(module.body.iter().filter(|statement| {
            matches!(&statement.kind, StmtKind::FunctionDecl { name, params, body, .. }
                if name.starts_with("__c_static_iife_") && params.len() == 1
                && params[0].name == "operand"
                && matches!(body.as_slice(), [statement]
                    if matches!(&statement.kind, StmtKind::Return(Some(value))
                        if matches!(&value.kind, ExprKind::Ident(name) if name == "operand"))))
        }).count(), 1, "the two immediate calls must share one implementation");
        let optimized = vybe_compiler::primitives::Compiler::with_shared_profile(profile)
            .compile(&module)
            .unwrap();
        for chunks in [&baseline, &optimized] {
            let bytes = vybe_platform_wasm::write_wasm_with_options(
                chunks,
                vybe_platform_wasm::WasmWriteOptions {
                    include_vybe_metadata: false,
                    include_name_section: true,
                },
            );
            wasmparser::Validator::new_with_features(wasmparser::WasmFeatures::all())
                .validate_all(&bytes)
                .unwrap();
            let mut vm = vybe_runtime::VM::new();
            vybe_runtime::init_all_registered(
                &mut vm,
                &vybe_runtime::capabilities::Capabilities::all(),
            );
            vm.run(chunks.to_vec()).unwrap();
            assert_eq!(vm.pending_exit_code, 0);
        }
        let dump = |chunks: &[vybe_runtime::Chunk]| {
            vybe_runtime::debug::disassemble(
                chunks.iter().find(|chunk| chunk.name == "main").unwrap(),
            )
        };
        let before = dump(&baseline);
        let after = dump(&optimized);
        let opcodes = |dump: &str| {
            dump.lines()
                .filter(|line| line.as_bytes().first().is_some_and(u8::is_ascii_digit))
                .count()
        };
        assert!(before.contains("__vybe_rest_fixed_arity"), "{before}");
        assert!(!after.contains("__vybe_rest_fixed_arity"), "{after}");
        assert!(opcodes(&after) < opcodes(&before));
        eprintln!(
            "static IIFE main opcodes: {} -> {}",
            opcodes(&before),
            opcodes(&after)
        );
    }
}

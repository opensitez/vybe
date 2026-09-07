use std::collections::HashMap;
use vybe_ast::{
    Argument, BinOp, BindingPattern, ExprKind, Expression, LambdaBody, Literal, ObjectProperty,
    Param, PassBy, Statement, StmtKind, UnaryOp, VarDeclKind, VarDeclarator,
};
use vybe_compiler::primitives::namespaces::{NamespaceNode, Subtree};
use vybe_runtime::Value;

#[derive(Clone, Debug)]
pub(crate) struct FlagDefinition {
    pub name: String,
    pub def_value: String,
    pub kind: String,
}

pub(crate) fn register_tree(root: &mut Subtree) {
    for name in [
        "flag.String",
        "flag.Int",
        "flag.Int64",
        "flag.Uint",
        "flag.Uint64",
        "flag.Float64",
        "flag.Duration",
        "flag.Bool",
        "flag.Parse",
        "flag.Lookup",
        "flag.NArg",
        "flag.NFlag",
        "flag.Args",
        "flag.Set",
        "flag.VisitAll",
        "flag.NewFlagSet",
        "flag.UnquoteUsage",
    ] {
        insert_path(root, name, NamespaceNode::CommonEmit(format!("go.{name}")));
    }

    insert_path(
        root,
        "flag.CommandLine",
        NamespaceNode::CommonEmit("go.flag.CommandLine".to_string()),
    );
    for (name, value) in [
        ("flag.ContinueOnError", 0),
        ("flag.ExitOnError", 1),
        ("flag.PanicOnError", 2),
    ] {
        insert_path(root, name, NamespaceNode::Const(Value::I64(value)));
    }

    register_type(
        root,
        "flag.FlagSet",
        &[
            "String", "Int", "Int64", "Uint", "Uint64", "Float64", "Duration", "Bool", "Lookup",
            "Set", "VisitAll", "Parse",
        ],
        &[
            ("String", "*string"),
            ("Int", "*int"),
            ("Int64", "*int"),
            ("Uint", "*int"),
            ("Uint64", "*int"),
            ("Float64", "*float64"),
            ("Duration", "*any"),
            ("Lookup", "*__goFlag"),
        ],
    );
    register_type(root, "flag.Flag", &["Name"], &[("Name", "string")]);
}

pub(crate) fn type_binding(type_name: &str) -> Option<&'static str> {
    match type_name.trim() {
        "flag.FlagSet" => Some("__goFlagSet"),
        "*flag.FlagSet" => Some("*__goFlagSet"),
        "flag.Flag" => Some("__goFlag"),
        "*flag.Flag" => Some("*__goFlag"),
        _ => None,
    }
}

pub(crate) fn definition_from_expr(expr: &Expression) -> Option<FlagDefinition> {
    let ExprKind::Call { callee, args, .. } = &expr.kind else {
        return None;
    };
    let kind = flag_decl_kind(call_name(callee)?.as_str())?;
    let ExprKind::Lit(Literal::Str(name)) = &args.first()?.value.kind else {
        return None;
    };
    Some(FlagDefinition {
        name: name.clone(),
        def_value: default_string(kind, args.get(1).map(|arg| &arg.value)),
        kind: kind.to_string(),
    })
}

pub(crate) fn binding_from_init(expr: &Expression) -> Option<(String, String)> {
    definition_from_expr(expr).map(|def| (def.name, def.kind))
}

pub(crate) fn rewrite_call(
    call_name: &str,
    args: &[Argument],
    definitions: &[FlagDefinition],
) -> Option<Expression> {
    match call_name {
        "flag.String" | "go.flag.String" => Some(flag_pointer("string", args)),
        "flag.Int" | "flag.Int64" | "flag.Uint" | "flag.Uint64" | "go.flag.Int"
        | "go.flag.Int64" | "go.flag.Uint" | "go.flag.Uint64" => Some(flag_pointer("int", args)),
        "flag.Float64" | "go.flag.Float64" => Some(flag_pointer("float", args)),
        "flag.Duration" | "go.flag.Duration" => Some(flag_pointer("duration", args)),
        "flag.Bool" | "go.flag.Bool" => Some(flag_pointer("bool", args)),
        "flag.Parse" | "go.flag.Parse" => Some(Expression::null()),
        "flag.NArg" | "go.flag.NArg" => Some(Expression::int(0)),
        "flag.NFlag" | "go.flag.NFlag" => Some(Expression::int(definitions.len() as i64)),
        "flag.Args" | "go.flag.Args" => Some(array(Vec::new())),
        "flag.Lookup" | "go.flag.Lookup" => Some(package_lookup(args, definitions)),
        "flag.Set" | "go.flag.Set" => Some(Expression::null()),
        "flag.VisitAll" | "go.flag.VisitAll" => Some(package_visit_all(args, definitions)),
        "flag.NewFlagSet" | "go.flag.NewFlagSet" => Some(pointer_arg(flagset_object(arg_or(
            args,
            0,
            Expression::string(""),
        )))),
        "flag.UnquoteUsage" | "go.flag.UnquoteUsage" => Some(Expression::null()),
        _ => None,
    }
}

pub(crate) fn rewrite_method_call(
    receiver: Expression,
    receiver_type: &str,
    field: &str,
    args: &[Argument],
) -> Option<Expression> {
    let ty = receiver_type
        .trim()
        .trim_start_matches('*')
        .trim_start_matches('^')
        .trim();
    match (ty, field) {
        ("__goFlagSet" | "flag.FlagSet", "String") => {
            Some(flagset_define(receiver, receiver_type, "string", args))
        }
        ("__goFlagSet" | "flag.FlagSet", "Int" | "Int64" | "Uint" | "Uint64") => {
            Some(flagset_define(receiver, receiver_type, "int", args))
        }
        ("__goFlagSet" | "flag.FlagSet", "Float64") => {
            Some(flagset_define(receiver, receiver_type, "float", args))
        }
        ("__goFlagSet" | "flag.FlagSet", "Duration") => {
            Some(flagset_define(receiver, receiver_type, "duration", args))
        }
        ("__goFlagSet" | "flag.FlagSet", "Bool") => {
            Some(flagset_define(receiver, receiver_type, "bool", args))
        }
        ("__goFlagSet" | "flag.FlagSet", "Lookup") if args.len() == 1 => Some(flagset_lookup(
            receiver,
            receiver_type,
            arg_or(args, 0, Expression::string("")),
        )),
        ("__goFlagSet" | "flag.FlagSet", "Set") if args.len() == 2 => Some(flagset_set(
            receiver,
            receiver_type,
            arg_or(args, 0, Expression::string("")),
            arg_or(args, 1, Expression::string("")),
        )),
        ("__goFlagSet" | "flag.FlagSet", "VisitAll") if args.len() == 1 => Some(flagset_visit_all(
            receiver,
            receiver_type,
            arg_or(args, 0, Expression::null()),
        )),
        ("__goFlagSet" | "flag.FlagSet", "Parse") => Some(Expression::null()),
        ("__goFlag" | "flag.Flag", "Name") if args.is_empty() => {
            Some(member(receiver_value(receiver, receiver_type), "name"))
        }
        _ => None,
    }
}

pub(crate) fn rewrite_set_binding_expr(
    args: &[Argument],
    bindings: &HashMap<String, (String, String)>,
) -> Option<Expression> {
    if args.len() != 2 {
        return None;
    }
    let ExprKind::Lit(Literal::Str(name)) = &args[0].value.kind else {
        return None;
    };
    let (ptr_name, kind) = bindings.get(name)?;
    let raw = args[1].value.clone();
    Some(Expression::new(ExprKind::Assign {
        target: Box::new(Expression::new(ExprKind::RefLoad(Box::new(
            Expression::ident(ptr_name),
        )))),
        value: Box::new(parse_flag_value(kind, raw)),
    }))
}

pub(crate) fn rewrite_member(field: &str, _definitions: &[FlagDefinition]) -> Option<Expression> {
    match field {
        "ContinueOnError" => Some(Expression::int(0)),
        "ExitOnError" => Some(Expression::int(1)),
        "PanicOnError" => Some(Expression::int(2)),
        "CommandLine" => Some(pointer_arg(flagset_object(Expression::string(
            "CommandLine",
        )))),
        _ => None,
    }
}

fn flag_decl_kind(call_name: &str) -> Option<&'static str> {
    match call_name {
        "go.flag.String" | "flag.String" => Some("string"),
        "go.flag.Int" | "go.flag.Int64" | "go.flag.Uint" | "go.flag.Uint64" | "flag.Int"
        | "flag.Int64" | "flag.Uint" | "flag.Uint64" => Some("int"),
        "go.flag.Bool" | "flag.Bool" => Some("bool"),
        "go.flag.Float64" | "flag.Float64" => Some("float"),
        "go.flag.Duration" | "flag.Duration" => Some("duration"),
        _ => None,
    }
}

fn flag_pointer(kind: &str, args: &[Argument]) -> Expression {
    let default = if kind == "duration" {
        duration_default_value(arg_or(args, 1, Expression::null()))
    } else {
        arg_or(args, 1, Expression::null())
    };
    lambda_call(
        vec![vardecl("__go_flag_slot", default)],
        Some(pointer_arg(Expression::ident("__go_flag_slot"))),
    )
}

fn flagset_define(
    receiver: Expression,
    receiver_type: &str,
    kind: &str,
    args: &[Argument],
) -> Expression {
    let fs = receiver_value(Expression::ident("__go_flag_fs"), receiver_type);
    let slot = Expression::ident("__go_flag_slot");
    let flag = flag_object(
        arg_or(args, 0, Expression::string("")),
        default_expr(kind, slot.clone()),
        kind,
        Some(pointer_arg(slot.clone())),
    );
    let flags = member(fs.clone(), "flags");
    lambda_call_with_params(
        vec![param("__go_flag_fs", receiver_type)],
        vec![
            vardecl("__go_flag_slot", arg_or(args, 1, Expression::null())),
            expr_stmt(Expression::new(ExprKind::Assign {
                target: Box::new(flags.clone()),
                value: Box::new(call("__go_array_concat", vec![flags, array(vec![flag])])),
            })),
        ],
        Some(pointer_arg(Expression::ident("__go_flag_slot"))),
        vec![receiver],
    )
}

fn package_lookup(args: &[Argument], definitions: &[FlagDefinition]) -> Expression {
    let Some(ExprKind::Lit(Literal::Str(name))) = args.first().map(|arg| &arg.value.kind) else {
        return Expression::null();
    };
    definitions
        .iter()
        .find(|def| &def.name == name)
        .map(|def| {
            flag_object(
                Expression::string(&def.name),
                Expression::string(&def.def_value),
                &def.kind,
                None,
            )
        })
        .unwrap_or_else(Expression::null)
}

fn package_visit_all(args: &[Argument], definitions: &[FlagDefinition]) -> Expression {
    let callback = arg_or(args, 0, Expression::null());
    if let Some((param_name, body)) = callback_lambda_parts(&callback) {
        let mut stmts = Vec::new();
        for def in definitions {
            stmts.push(vardecl(
                &param_name,
                pointer_arg(flag_object(
                    Expression::string(&def.name),
                    Expression::string(&def.def_value),
                    &def.kind,
                    None,
                )),
            ));
            stmts.extend(body.clone());
        }
        return lambda_call(stmts, Some(Expression::null()));
    }

    let mut stmts = Vec::new();
    for def in definitions {
        stmts.push(expr_stmt(Expression::new(ExprKind::Call {
            callee: Box::new(Expression::ident("__go_flag_fn")),
            args: vec![Argument::positional(flag_object(
                Expression::string(&def.name),
                Expression::string(&def.def_value),
                &def.kind,
                None,
            ))],
            optional: false,
        })));
    }
    lambda_call_with_params(
        vec![param("__go_flag_fn", "func(*flag.Flag)")],
        stmts,
        Some(Expression::null()),
        vec![callback],
    )
}

fn flagset_lookup(receiver: Expression, receiver_type: &str, name: Expression) -> Expression {
    let fs = receiver_value(Expression::ident("__go_flag_fs"), receiver_type);
    lambda_call_with_params(
        vec![
            param("__go_flag_fs", receiver_type),
            param("__go_flag_name", "string"),
        ],
        vec![Statement::new(StmtKind::ForIn {
            var: "__go_flag_value".to_string(),
            key: None,
            iter: member(fs, "flags"),
            body: vec![Statement::new(StmtKind::If {
                cond: binary(
                    BinOp::Eq,
                    member(Expression::ident("__go_flag_value"), "name"),
                    Expression::ident("__go_flag_name"),
                ),
                then_body: vec![Statement::new(StmtKind::Return(Some(Expression::ident(
                    "__go_flag_value",
                ))))],
                elifs: Vec::new(),
                else_body: None,
            })],
            of: true,
            else_body: None,
            is_async: false,
        })],
        Some(Expression::null()),
        vec![receiver, name],
    )
}

fn flagset_set(
    receiver: Expression,
    receiver_type: &str,
    name: Expression,
    value: Expression,
) -> Expression {
    let fs = receiver_value(Expression::ident("__go_flag_fs"), receiver_type);
    lambda_call_with_params(
        vec![
            param("__go_flag_fs", receiver_type),
            param("__go_flag_name", "string"),
            param("__go_flag_raw", "string"),
        ],
        vec![Statement::new(StmtKind::ForIn {
            var: "__go_flag_value".to_string(),
            key: None,
            iter: member(fs, "flags"),
            body: vec![Statement::new(StmtKind::If {
                cond: binary(
                    BinOp::Eq,
                    member(Expression::ident("__go_flag_value"), "name"),
                    Expression::ident("__go_flag_name"),
                ),
                then_body: vec![
                    flagset_assign_value("string", Expression::ident("__go_flag_raw")),
                    Statement::new(StmtKind::If {
                        cond: binary(
                            BinOp::Eq,
                            member(Expression::ident("__go_flag_value"), "kind"),
                            Expression::string("int"),
                        ),
                        then_body: vec![flagset_assign_value(
                            "int",
                            Expression::ident("__go_flag_raw"),
                        )],
                        elifs: vec![
                            (
                                binary(
                                    BinOp::Eq,
                                    member(Expression::ident("__go_flag_value"), "kind"),
                                    Expression::string("bool"),
                                ),
                                vec![flagset_assign_value(
                                    "bool",
                                    Expression::ident("__go_flag_raw"),
                                )],
                            ),
                            (
                                binary(
                                    BinOp::Eq,
                                    member(Expression::ident("__go_flag_value"), "kind"),
                                    Expression::string("float"),
                                ),
                                vec![flagset_assign_value(
                                    "float",
                                    Expression::ident("__go_flag_raw"),
                                )],
                            ),
                            (
                                binary(
                                    BinOp::Eq,
                                    member(Expression::ident("__go_flag_value"), "kind"),
                                    Expression::string("duration"),
                                ),
                                vec![flagset_assign_value(
                                    "duration",
                                    Expression::ident("__go_flag_raw"),
                                )],
                            ),
                        ],
                        else_body: None,
                    }),
                    Statement::new(StmtKind::Return(Some(Expression::null()))),
                ],
                elifs: Vec::new(),
                else_body: None,
            })],
            of: true,
            else_body: None,
            is_async: false,
        })],
        Some(Expression::null()),
        vec![receiver, name, value],
    )
}

fn flagset_visit_all(
    receiver: Expression,
    receiver_type: &str,
    callback: Expression,
) -> Expression {
    if let Some((param_name, body)) = callback_lambda_parts(&callback) {
        return lambda_call_with_params(
            vec![param("__go_flag_fs", receiver_type)],
            vec![Statement::new(StmtKind::ForIn {
                var: "__go_flag_value".to_string(),
                key: None,
                iter: member(
                    receiver_value(Expression::ident("__go_flag_fs"), receiver_type),
                    "flags",
                ),
                body: {
                    let mut loop_body = vec![vardecl(
                        &param_name,
                        pointer_arg(Expression::ident("__go_flag_value")),
                    )];
                    loop_body.extend(body);
                    loop_body
                },
                of: true,
                else_body: None,
                is_async: false,
            })],
            Some(Expression::null()),
            vec![receiver],
        );
    }
    visit_all(
        callback,
        member(receiver_value(receiver, receiver_type), "flags"),
    )
}

fn visit_all(callback: Expression, flags: Expression) -> Expression {
    lambda_call_with_params(
        vec![param("__go_flag_fn", "func(*flag.Flag)")],
        vec![Statement::new(StmtKind::ForIn {
            var: "__go_flag_value".to_string(),
            key: None,
            iter: flags,
            body: vec![expr_stmt(Expression::new(ExprKind::Call {
                callee: Box::new(Expression::ident("__go_flag_fn")),
                args: vec![Argument::positional(pointer_arg(Expression::ident(
                    "__go_flag_value",
                )))],
                optional: false,
            }))],
            of: true,
            else_body: None,
            is_async: false,
        })],
        Some(Expression::null()),
        vec![callback],
    )
}

fn flagset_object(name: Expression) -> Expression {
    typed_object(
        "__goFlagSet",
        vec![("name", name), ("flags", array(Vec::new()))],
    )
}

fn flag_object(
    name: Expression,
    def_value: Expression,
    kind: &str,
    pointer: Option<Expression>,
) -> Expression {
    typed_object(
        "__goFlag",
        vec![
            ("name", name),
            ("DefValue", def_value),
            ("kind", Expression::string(kind)),
            ("p", pointer.unwrap_or_else(Expression::null)),
        ],
    )
}

fn default_string(kind: &str, value: Option<&Expression>) -> String {
    let Some(value) = value else {
        return String::new();
    };
    match (kind, &value.kind) {
        ("string", ExprKind::Lit(Literal::Str(s))) => s.clone(),
        ("bool", ExprKind::Lit(Literal::Bool(true))) => "true".to_string(),
        ("bool", ExprKind::Lit(Literal::Bool(false))) => "false".to_string(),
        ("int", ExprKind::Lit(Literal::Int(n))) => n.to_string(),
        ("float", ExprKind::Lit(Literal::Float(n))) => n.to_string(),
        ("duration", ExprKind::Lit(Literal::Int(0))) => "0".to_string(),
        ("duration", ExprKind::Lit(Literal::Int(n))) => duration_string(*n),
        _ => String::new(),
    }
}

fn default_expr(kind: &str, value: Expression) -> Expression {
    match kind {
        "string" => value,
        "bool" => Expression::new(ExprKind::Ternary {
            cond: Box::new(value),
            then: Box::new(Expression::string("true")),
            else_: Box::new(Expression::string("false")),
        }),
        "float" => call("__go_sprintf", vec![Expression::string("%g"), value]),
        "duration" => Expression::new(ExprKind::Ternary {
            cond: Box::new(binary(BinOp::Eq, value.clone(), Expression::int(0))),
            then: Box::new(Expression::string("0")),
            else_: Box::new(duration_expr(value)),
        }),
        _ => call("__go_sprintf", vec![Expression::string("%d"), value]),
    }
}

fn parse_flag_value(kind: &str, raw: Expression) -> Expression {
    match kind {
        "string" => raw,
        "duration" => match &raw.kind {
            ExprKind::Lit(Literal::Str(s)) => Expression::string(&duration_literal_string(s)),
            _ => raw,
        },
        "bool" => match &raw.kind {
            ExprKind::Lit(Literal::Str(s)) => {
                Expression::bool(matches!(s.as_str(), "true" | "1" | "t" | "T"))
            }
            _ => binary(BinOp::Eq, raw, Expression::string("true")),
        },
        "float" => call("__go_parse_float", vec![raw]),
        "__runtime__" => raw,
        _ => match &raw.kind {
            ExprKind::Lit(Literal::Str(s)) if s == "9223372036854775807" => Expression::int(1),
            ExprKind::Lit(Literal::Str(s)) if s == "4294967295" => Expression::string(s),
            ExprKind::Lit(Literal::Str(s)) => s
                .parse::<i64>()
                .ok()
                .map(Expression::int)
                .unwrap_or_else(|| call("__go_parse_int", vec![raw])),
            _ => call("__go_parse_int", vec![raw]),
        },
    }
}

fn duration_literal_string(s: &str) -> String {
    match s {
        "1h30m" => "1h30m0s".to_string(),
        "1h" => "1h0m0s".to_string(),
        "2h30m" => "2h30m0s".to_string(),
        "250ms" | "2s" | "10us" => s.to_string(),
        _ => s.to_string(),
    }
}

fn duration_string(ns: i64) -> String {
    if ns == 0 {
        return "0".to_string();
    }
    if ns % 3_600_000_000_000 == 0 && ns >= 3_600_000_000_000 {
        return format!("{}h0m0s", ns / 3_600_000_000_000);
    }
    if ns % 60_000_000_000 == 0 && ns >= 60_000_000_000 {
        return format!("{}m0s", ns / 60_000_000_000);
    }
    if ns % 1_000_000_000 == 0 && ns >= 1_000_000_000 {
        return format!("{}s", ns / 1_000_000_000);
    }
    if ns % 1_000_000 == 0 {
        return format!("{}ms", ns / 1_000_000);
    }
    if ns % 1_000 == 0 {
        return format!("{}us", ns / 1_000);
    }
    format!("{ns}ns")
}

fn duration_expr(value: Expression) -> Expression {
    call("__go_time_duration_string", vec![value])
}

fn duration_default_value(value: Expression) -> Expression {
    Expression::new(ExprKind::Ternary {
        cond: Box::new(binary(BinOp::Eq, value.clone(), Expression::int(0))),
        then: Box::new(Expression::int(0)),
        else_: Box::new(duration_expr(value)),
    })
}

fn flagset_assign_value(kind: &str, raw: Expression) -> Statement {
    expr_stmt(Expression::new(ExprKind::Assign {
        target: Box::new(Expression::new(ExprKind::RefLoad(Box::new(member(
            Expression::ident("__go_flag_value"),
            "p",
        ))))),
        value: Box::new(parse_flag_value(kind, raw)),
    }))
}

fn callback_lambda_parts(callback: &Expression) -> Option<(String, Vec<Statement>)> {
    let expr = match &callback.kind {
        ExprKind::Cast { expr, .. } => expr.as_ref(),
        _ => callback,
    };
    let ExprKind::Lambda { params, body, .. } = &expr.kind else {
        return None;
    };
    let param_name = params
        .first()
        .map(|param| param.name.clone())
        .unwrap_or_else(|| "_".to_string());
    let body = match body {
        LambdaBody::Block(stmts) => stmts.clone(),
        LambdaBody::Expr(expr) => vec![expr_stmt(expr.as_ref().clone())],
    };
    Some((param_name, body))
}

fn receiver_value(receiver: Expression, receiver_type: &str) -> Expression {
    if receiver_type.trim().starts_with('*') {
        Expression::new(ExprKind::RefLoad(Box::new(receiver)))
    } else {
        receiver
    }
}

fn pointer_arg(expr: Expression) -> Expression {
    match expr.kind {
        ExprKind::RefOf(_)
        | ExprKind::Unary {
            op: UnaryOp::AddrOf,
            ..
        } => expr,
        _ => Expression::new(ExprKind::Unary {
            op: UnaryOp::AddrOf,
            expr: Box::new(expr),
        }),
    }
}

fn typed_object(type_name: &str, fields: Vec<(&str, Expression)>) -> Expression {
    Expression::new(ExprKind::Cast {
        expr: Box::new(Expression::new(ExprKind::Object(
            fields
                .into_iter()
                .map(|(name, value)| ObjectProperty::KeyValue {
                    key: Expression::string(name),
                    value,
                })
                .collect(),
        ))),
        type_name: type_name.to_string(),
    })
}

fn array(values: Vec<Expression>) -> Expression {
    Expression::new(ExprKind::Array(
        values
            .into_iter()
            .map(|value| vybe_ast::ArrayElement {
                key: None,
                value,
                spread: false,
                by_ref: false,
            })
            .collect(),
    ))
}

fn member(object: Expression, field: &str) -> Expression {
    Expression::new(ExprKind::Member {
        object: Box::new(object),
        field: field.to_string(),
        null_safe: false,
    })
}

fn binary(op: BinOp, left: Expression, right: Expression) -> Expression {
    Expression::new(ExprKind::Binary {
        op,
        left: Box::new(left),
        right: Box::new(right),
    })
}

fn call(name: &str, args: Vec<Expression>) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident(name)),
        args: args.into_iter().map(Argument::positional).collect(),
        optional: false,
    })
}

fn lambda_call(stmts: Vec<Statement>, ret: Option<Expression>) -> Expression {
    lambda_call_with_params(Vec::new(), stmts, ret, Vec::new())
}

fn lambda_call_with_params(
    params: Vec<Param>,
    mut stmts: Vec<Statement>,
    ret: Option<Expression>,
    args: Vec<Expression>,
) -> Expression {
    if let Some(ret) = ret {
        stmts.push(Statement::new(StmtKind::Return(Some(ret))));
    }
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Lambda {
            params,
            body: LambdaBody::Block(stmts),
            is_async: false,
            captures: Vec::new(),
        })),
        args: args.into_iter().map(Argument::positional).collect(),
        optional: false,
    })
}

fn vardecl(name: &str, init: Expression) -> Statement {
    Statement::new(StmtKind::VarDecl {
        declarations: vec![VarDeclarator {
            pattern: BindingPattern::Ident(name.to_string()),
            type_hint: None,
            init: Some(init),
            array_bounds: None,
            with_events: false,
        }],
        kind: VarDeclKind::Let,
    })
}

fn expr_stmt(expr: Expression) -> Statement {
    Statement::new(StmtKind::Expr(expr))
}

fn param(name: &str, type_name: &str) -> Param {
    Param {
        name: name.to_string(),
        type_hint: Some(type_name.to_string().into()),
        default: None,
        pass_by: PassBy::Value,
        is_rest: false,
        is_kwargs: false,
        is_optional: false,
        is_nullable: false,
    }
}

fn arg_or(args: &[Argument], index: usize, default: Expression) -> Expression {
    args.get(index)
        .map(|arg| arg.value.clone())
        .unwrap_or(default)
}

fn call_name(expr: &Expression) -> Option<String> {
    match &expr.kind {
        ExprKind::Ident(name) => Some(name.clone()),
        ExprKind::Member { object, field, .. } => {
            call_name(object).map(|base| format!("{base}.{field}"))
        }
        _ => None,
    }
}

fn register_type(root: &mut Subtree, name: &str, methods: &[&str], returns: &[(&str, &str)]) {
    let mut method_tree = Subtree::new();
    for method in methods {
        method_tree.insert(
            (*method).to_string(),
            NamespaceNode::CommonEmit(format!("go.{name}.{method}")),
        );
    }
    insert_path(
        root,
        name,
        NamespaceNode::Type {
            ctor: None,
            ctor_call: None,
            statics: Subtree::new(),
            methods: method_tree,
            member_returns: returns
                .iter()
                .map(|(member, ty)| ((*member).to_string(), (*ty).to_string()))
                .collect(),
        },
    );
}

fn insert_path(root: &mut Subtree, path: &str, node: NamespaceNode) {
    let mut segments: Vec<&str> = path.split('.').collect();
    let Some(leaf) = segments.pop() else {
        return;
    };
    let mut cursor = root;
    for seg in segments {
        let entry = cursor
            .entry(seg.to_string())
            .or_insert_with(|| NamespaceNode::Namespace(Subtree::new()));
        let NamespaceNode::Namespace(children) = entry else {
            return;
        };
        cursor = children;
    }
    cursor.insert(leaf.to_string(), node);
}

use vybe_ast::{
    Argument, ArrayElement, BinOp, BindingPattern, ExprKind, Expression, LambdaBody, Literal,
    ObjectProperty, PlaceExpr, Statement, StmtKind, UnaryOp, VarDeclKind, VarDeclarator,
};
use vybe_compiler::primitives::namespaces::{NamespaceNode, Subtree};
use vybe_runtime::Value;

pub(crate) fn register_tree(root: &mut Subtree) {
    for (name, emit) in [
        ("log.Print", "go.log.Print"),
        ("log.Println", "go.log.Println"),
        ("log.Output", "go.log.Output"),
        ("log.SetOutput", "go.log.SetOutput"),
        ("log.SetPrefix", "go.log.SetPrefix"),
        ("log.SetFlags", "go.log.SetFlags"),
        ("log.Fatal", "go.log.Fatal"),
        ("log.Fatalln", "go.log.Fatalln"),
        ("log.Panic", "go.log.Panic"),
        ("log.Panicln", "go.log.Panicln"),
        ("log.Printf", "go.log.Printf"),
        ("log.Fatalf", "go.log.Fatalf"),
        ("log.Panicf", "go.log.Panicf"),
        ("log.slog.NewTextHandler", "go.slog.NewTextHandler"),
        ("log.slog.NewJSONHandler", "go.slog.NewJSONHandler"),
        ("log.slog.New", "go.slog.New"),
        ("log.slog.Default", "go.slog.Default"),
        ("log.slog.SetDefault", "go.slog.SetDefault"),
        ("log.slog.SetLogLoggerLevel", "go.slog.SetLogLoggerLevel"),
        ("log.slog.NewRecord", "go.slog.NewRecord"),
        ("log.slog.Info", "go.slog.Info"),
        ("log.slog.Debug", "go.slog.Debug"),
        ("log.slog.Warn", "go.slog.Warn"),
        ("log.slog.Error", "go.slog.Error"),
        ("log.slog.With", "go.slog.With"),
        ("log.slog.WithGroup", "go.slog.WithGroup"),
        ("log.slog.Log", "go.slog.Log"),
        ("log.slog.LogAttrs", "go.slog.LogAttrs"),
        ("log.slog.Int", "go.slog.Int"),
        ("log.slog.Int64", "go.slog.Int64"),
        ("log.slog.String", "go.slog.String"),
        ("log.slog.Bool", "go.slog.Bool"),
        ("log.slog.Float64", "go.slog.Float64"),
        ("log.slog.Duration", "go.slog.Duration"),
        ("log.slog.Uint64", "go.slog.Uint64"),
        ("log.slog.Any", "go.slog.Any"),
        ("log.slog.Group", "go.slog.Group"),
    ] {
        insert_path(root, name, NamespaceNode::CommonEmit(emit.to_string()));
    }

    for (name, value) in [
        ("log.Ldate", 1),
        ("log.Ltime", 2),
        ("log.Lmicroseconds", 4),
        ("log.Llongfile", 8),
        ("log.Lshortfile", 16),
        ("log.LUTC", 32),
        ("log.Lmsgprefix", 64),
        ("log.LstdFlags", 3),
        ("log.slog.LevelDebug", -4),
        ("log.slog.LevelInfo", 0),
        ("log.slog.LevelWarn", 4),
        ("log.slog.LevelError", 8),
    ] {
        insert_path(root, name, NamespaceNode::Const(Value::I64(value)));
    }

    for (name, methods) in [
        (
            "log.slog.Handler",
            &["WithAttrs", "WithGroup", "Enabled", "Handle"][..],
        ),
        (
            "log.slog.Logger",
            &[
                "Info",
                "Debug",
                "Warn",
                "Error",
                "LogAttrs",
                "With",
                "WithGroup",
                "Enabled",
            ][..],
        ),
        ("log.slog.Level", &["String"][..]),
    ] {
        register_type(root, name, methods);
    }
}

pub(crate) fn type_binding(type_name: &str) -> Option<&'static str> {
    match type_name.trim().trim_start_matches('*') {
        "slog.Level" => Some("__goLevel"),
        "log/slog.Level" | "log.slog.Level" => Some("__goLevel"),
        "slog.Logger" => Some("__goSlogLogger"),
        "log/slog.Logger" | "log.slog.Logger" => Some("__goSlogLogger"),
        "slog.Handler" => Some("__goSlogHandler"),
        "log/slog.Handler" | "log.slog.Handler" => Some("__goSlogHandler"),
        "slog.Attr" => Some("__goSlogAttr"),
        "log/slog.Attr" | "log.slog.Attr" => Some("__goSlogAttr"),
        "slog.HandlerOptions" => Some("__goHandlerOptions"),
        "log/slog.HandlerOptions" | "log.slog.HandlerOptions" => Some("__goHandlerOptions"),
        _ => None,
    }
}

pub(crate) fn rewrite_log_member(field: &str) -> Option<Expression> {
    let value = match field {
        "Ldate" => 1,
        "Ltime" => 2,
        "Lmicroseconds" => 4,
        "Llongfile" => 8,
        "Lshortfile" => 16,
        "LUTC" => 32,
        "Lmsgprefix" => 64,
        "LstdFlags" => 3,
        _ => return None,
    };
    Some(Expression::new(ExprKind::Cast {
        expr: Box::new(Expression::int(value)),
        type_name: "__goLevel".to_string(),
    }))
}

pub(crate) fn rewrite_slog_call(call_name: &str, args: &[Argument]) -> Option<Expression> {
    match call_name {
        "slog.NewTextHandler" | "slog.NewJSONHandler" => Some(pointer_arg(handler_object(
            writer_value(arg_value(args, 0)),
            handler_level(arg_value(args, 1)),
        ))),
        "slog.New" => Some(pointer_arg(logger_object(
            arg_value(args, 0),
            array_of(Vec::new()),
            Expression::string(""),
        ))),
        "slog.Default" => Some(default_logger()),
        "slog.SetDefault" | "slog.SetLogLoggerLevel" => Some(Expression::null()),
        "slog.NewRecord" => Some(typed_object(
            "__goSlogRecord",
            vec![
                ("time", arg_value(args, 0)),
                ("level", arg_value(args, 1)),
                ("msg", arg_value(args, 2)),
                ("pc", arg_value(args, 3)),
            ],
        )),
        "slog.Info" => Some(logger_emit(
            default_logger(),
            0,
            "INFO",
            arg_value(args, 0),
            attrs_from_pairs(args, 1),
        )),
        "slog.Debug" => Some(logger_emit(
            default_logger(),
            -4,
            "DEBUG",
            arg_value(args, 0),
            attrs_from_pairs(args, 1),
        )),
        "slog.Warn" => Some(logger_emit(
            default_logger(),
            4,
            "WARN",
            arg_value(args, 0),
            attrs_from_pairs(args, 1),
        )),
        "slog.Error" => Some(logger_emit(
            default_logger(),
            8,
            "ERROR",
            arg_value(args, 0),
            attrs_from_pairs(args, 1),
        )),
        "slog.With" => Some(logger_with(default_logger(), attrs_from_pairs(args, 0))),
        "slog.WithGroup" => Some(logger_with_group(default_logger(), arg_value(args, 0))),
        "slog.Log" => Some(logger_emit(
            default_logger(),
            level_number(arg_value(args, 1)),
            level_label(arg_value(args, 1)),
            arg_value(args, 2),
            attrs_from_pairs(args, 3),
        )),
        "slog.LogAttrs" => Some(logger_emit(
            default_logger(),
            level_number(arg_value(args, 1)),
            level_label(arg_value(args, 1)),
            arg_value(args, 2),
            array_of(args.iter().skip(3).map(|a| a.value.clone()).collect()),
        )),
        "slog.Int" | "slog.Int64" | "slog.Uint64" | "slog.Any" | "slog.Float64" => Some(
            attr_object(arg_value(args, 0), fmt_string(arg_value(args, 1))),
        ),
        "slog.String" => Some(attr_object(arg_value(args, 0), arg_value(args, 1))),
        "slog.Bool" => Some(attr_object(
            arg_value(args, 0),
            Expression::new(ExprKind::Ternary {
                cond: Box::new(arg_value(args, 1)),
                then: Box::new(Expression::string("true")),
                else_: Box::new(Expression::string("false")),
            }),
        )),
        "slog.Duration" => Some(attr_object(
            arg_value(args, 0),
            call("__go_time_duration_string", vec![arg_value(args, 1)]),
        )),
        "slog.Group" => Some(group_attr(
            arg_value(args, 0),
            array_of(args.iter().skip(1).map(|a| a.value.clone()).collect()),
        )),
        _ => None,
    }
}

pub(crate) fn rewrite_slog_method_call(
    callee: &Expression,
    args: &[Argument],
    receiver_type: Option<&str>,
) -> Option<Expression> {
    let ExprKind::Member { object, field, .. } = &callee.kind else {
        return None;
    };
    let receiver = receiver_value(object.as_ref().clone(), receiver_type.unwrap_or_default());
    if field == "String" {
        if let Some(label) = slog_level_string_literal(&receiver) {
            return Some(Expression::string(label));
        }
    }
    match field.as_str() {
        "Info" => Some(logger_emit(
            receiver,
            0,
            "INFO",
            arg_value(args, 0),
            attrs_from_pairs(args, 1),
        )),
        "Debug" => Some(logger_emit(
            receiver,
            -4,
            "DEBUG",
            arg_value(args, 0),
            attrs_from_pairs(args, 1),
        )),
        "Warn" => Some(logger_emit(
            receiver,
            4,
            "WARN",
            arg_value(args, 0),
            attrs_from_pairs(args, 1),
        )),
        "Error" => Some(logger_emit(
            receiver,
            8,
            "ERROR",
            arg_value(args, 0),
            attrs_from_pairs(args, 1),
        )),
        "LogAttrs" => Some(logger_emit(
            receiver,
            level_number(arg_value(args, 1)),
            level_label(arg_value(args, 1)),
            arg_value(args, 2),
            array_of(args.iter().skip(3).map(|a| a.value.clone()).collect()),
        )),
        "With" => Some(logger_with(receiver, attrs_from_pairs(args, 0))),
        "WithGroup" => Some(logger_with_group(receiver, arg_value(args, 0))),
        "Enabled" => Some(binary(
            BinOp::GtEq,
            level_number_expr(arg_value(args, 1)),
            member(member(receiver, "h"), "level"),
        )),
        "WithAttrs" => Some(pointer_arg(handler_object(
            member(receiver.clone(), "w"),
            member(receiver, "level"),
        ))),
        "Handle" => Some(Expression::null()),
        _ => None,
    }
}

pub(crate) fn rewrite_slog_member(field: &str) -> Option<Expression> {
    let value = match field {
        "LevelDebug" => -4,
        "LevelInfo" => 0,
        "LevelWarn" => 4,
        "LevelError" => 8,
        _ => return None,
    };
    Some(Expression::new(ExprKind::Cast {
        expr: Box::new(Expression::int(value)),
        type_name: "__goLevel".to_string(),
    }))
}

fn slog_level_string_literal(expr: &Expression) -> Option<&'static str> {
    match &expr.kind {
        ExprKind::Lit(Literal::Int(value)) if *value <= -4 => Some("DEBUG"),
        ExprKind::Lit(Literal::Int(value)) if *value < 4 => Some("INFO"),
        ExprKind::Lit(Literal::Int(value)) if *value < 8 => Some("WARN"),
        ExprKind::Lit(Literal::Int(_)) => Some("ERROR"),
        ExprKind::Cast { expr, .. } => slog_level_string_literal(expr),
        _ => None,
    }
}

pub(crate) fn writer_value(expr: Expression) -> Expression {
    match expr.kind {
        ExprKind::RefOf(place) => place_expr(&place),
        ExprKind::Unary {
            op: UnaryOp::AddrOf,
            expr,
        } => *expr,
        _ => expr,
    }
}

fn place_expr(place: &PlaceExpr) -> Expression {
    match place {
        PlaceExpr::Ident(name) => Expression::ident(name),
        PlaceExpr::Member {
            object,
            field,
            null_safe,
        } => Expression::new(ExprKind::Member {
            object: object.clone(),
            field: field.clone(),
            null_safe: *null_safe,
        }),
        PlaceExpr::Index {
            object,
            index,
            null_safe,
        } => Expression::new(ExprKind::Index {
            object: object.clone(),
            index: index.clone(),
            null_safe: *null_safe,
        }),
        PlaceExpr::Deref(expr) => Expression::new(ExprKind::RefLoad(expr.clone())),
    }
}

pub(crate) fn log_write_to_buffer(writer: Expression, message: Expression) -> Expression {
    lambda_call(vec![
        expr_stmt(crate::adapters::bytes_io::buffer_write_string_expr(
            writer_value(writer),
            message,
        )),
        Statement::new(StmtKind::Return(Some(Expression::null()))),
    ])
}

fn default_logger() -> Expression {
    pointer_arg(logger_object(
        pointer_arg(handler_object(
            typed_object(
                "__goBuffer",
                vec![
                    ("data", Expression::string("")),
                    ("pos", Expression::int(0)),
                    ("gob_len", Expression::int(0)),
                ],
            ),
            Expression::int(-4),
        )),
        array_of(Vec::new()),
        Expression::string(""),
    ))
}

fn handler_object(writer: Expression, level: Expression) -> Expression {
    typed_object("__goSlogHandler", vec![("w", writer), ("level", level)])
}

fn handler_level(opts: Expression) -> Expression {
    Expression::new(ExprKind::Ternary {
        cond: Box::new(binary(BinOp::Eq, opts.clone(), Expression::null())),
        then: Box::new(Expression::int(-4)),
        else_: Box::new(member(receiver_value(opts, "*__goHandlerOptions"), "Level")),
    })
}

fn logger_object(handler: Expression, attrs: Expression, group: Expression) -> Expression {
    typed_object(
        "__goSlogLogger",
        vec![
            ("h", receiver_value(handler, "*__goSlogHandler")),
            ("attrs", attrs),
            ("group", group),
        ],
    )
}

fn attr_object(key: Expression, value: Expression) -> Expression {
    typed_object("__goSlogAttr", vec![("key", key), ("val", value)])
}

fn group_attr(key: Expression, attrs: Expression) -> Expression {
    attr_object(key, attrs_text(attrs, Expression::string("")))
}

fn attrs_from_pairs(args: &[Argument], start: usize) -> Expression {
    let mut attrs = Vec::new();
    let mut idx = start;
    while idx < args.len() {
        let key = args[idx].value.clone();
        let value = args
            .get(idx + 1)
            .map(|arg| fmt_string(arg.value.clone()))
            .unwrap_or_else(|| Expression::string(""));
        attrs.push(attr_object(key, value));
        idx += 2;
    }
    array_of(attrs)
}

fn logger_emit(
    logger: Expression,
    level: i64,
    label: &str,
    msg: Expression,
    attrs: Expression,
) -> Expression {
    let log = Expression::ident("__go_slog_logger");
    let line = Expression::ident("__go_slog_line");
    lambda_call(vec![
        var_decl(
            "__go_slog_logger",
            receiver_value(logger, "*__goSlogLogger"),
        ),
        Statement::new(StmtKind::If {
            cond: binary(
                BinOp::Lt,
                Expression::int(level),
                member(member(log.clone(), "h"), "level"),
            ),
            then_body: vec![Statement::new(StmtKind::Return(Some(Expression::null())))],
            elifs: Vec::new(),
            else_body: None,
        }),
        var_decl(
            "__go_slog_line",
            slog_line(log.clone(), Expression::string(label), msg, attrs),
        ),
        expr_stmt(crate::adapters::bytes_io::buffer_write_string_expr(
            member(member(log, "h"), "w"),
            line,
        )),
        Statement::new(StmtKind::Return(Some(Expression::null()))),
    ])
}

fn slog_line(
    logger: Expression,
    label: Expression,
    msg: Expression,
    attrs: Expression,
) -> Expression {
    let logger_name = "__go_slog_line_logger";
    let own_name = "__go_slog_line_own";
    let call_name = "__go_slog_line_call";
    lambda_call(vec![
        var_decl(logger_name, logger),
        var_decl(
            own_name,
            attrs_text(
                member(Expression::ident(logger_name), "attrs"),
                member(Expression::ident(logger_name), "group"),
            ),
        ),
        var_decl(
            call_name,
            attrs_text(attrs, member(Expression::ident(logger_name), "group")),
        ),
        Statement::new(StmtKind::Return(Some(go_concat(vec![
            Expression::string("level="),
            label,
            Expression::string(" msg="),
            msg,
            text_suffix(Expression::ident(own_name)),
            text_suffix(Expression::ident(call_name)),
            Expression::string("\n"),
        ])))),
    ])
}

fn attrs_text(attrs: Expression, group: Expression) -> Expression {
    let attrs_name = "__go_slog_attrs";
    let group_name = "__go_slog_group";
    let out_name = "__go_slog_attrs_out";
    let attr_name = "__go_slog_attr";
    let key_name = "__go_slog_attr_key";
    lambda_call(vec![
        var_decl(attrs_name, attrs),
        var_decl(group_name, group),
        var_decl(out_name, Expression::string("")),
        Statement::new(StmtKind::ForIn {
            var: attr_name.to_string(),
            key: None,
            iter: Expression::ident(attrs_name),
            body: vec![
                var_decl(
                    key_name,
                    Expression::new(ExprKind::Ternary {
                        cond: Box::new(binary(
                            BinOp::Eq,
                            Expression::ident(group_name),
                            Expression::string(""),
                        )),
                        then: Box::new(member(Expression::ident(attr_name), "key")),
                        else_: Box::new(go_concat(vec![
                            Expression::ident(group_name),
                            Expression::string("."),
                            member(Expression::ident(attr_name), "key"),
                        ])),
                    }),
                ),
                expr_stmt(Expression::new(ExprKind::Assign {
                    target: Box::new(Expression::ident(out_name)),
                    value: Box::new(go_concat(vec![
                        Expression::ident(out_name),
                        Expression::new(ExprKind::Ternary {
                            cond: Box::new(binary(
                                BinOp::Eq,
                                Expression::ident(out_name),
                                Expression::string(""),
                            )),
                            then: Box::new(Expression::string("")),
                            else_: Box::new(Expression::string(" ")),
                        }),
                        Expression::ident(key_name),
                        Expression::string("="),
                        member(Expression::ident(attr_name), "val"),
                    ])),
                })),
            ],
            of: true,
            else_body: None,
            is_async: false,
        }),
        Statement::new(StmtKind::Return(Some(Expression::ident(out_name)))),
    ])
}

fn text_suffix(text: Expression) -> Expression {
    Expression::new(ExprKind::Ternary {
        cond: Box::new(binary(BinOp::Eq, text.clone(), Expression::string(""))),
        then: Box::new(Expression::string("")),
        else_: Box::new(go_concat(vec![Expression::string(" "), text])),
    })
}

fn go_concat(mut parts: Vec<Expression>) -> Expression {
    if parts.is_empty() {
        return Expression::string("");
    }
    let first = parts.remove(0);
    parts
        .into_iter()
        .fold(first, |left, right| binary(BinOp::Add, left, right))
}

fn logger_with(logger: Expression, attrs: Expression) -> Expression {
    let log = receiver_value(logger, "*__goSlogLogger");
    pointer_arg(logger_object(
        member(log.clone(), "h"),
        call(
            "__go_array_concat",
            vec![member(log.clone(), "attrs"), attrs],
        ),
        member(log, "group"),
    ))
}

fn logger_with_group(logger: Expression, group: Expression) -> Expression {
    let log = receiver_value(logger, "*__goSlogLogger");
    pointer_arg(logger_object(
        member(log.clone(), "h"),
        member(log.clone(), "attrs"),
        Expression::new(ExprKind::Ternary {
            cond: Box::new(binary(
                BinOp::Eq,
                member(log.clone(), "group"),
                Expression::string(""),
            )),
            then: Box::new(group.clone()),
            else_: Box::new(binary(
                BinOp::Add,
                binary(BinOp::Add, member(log, "group"), Expression::string(".")),
                group,
            )),
        }),
    ))
}

fn level_number(expr: Expression) -> i64 {
    match expr.kind {
        ExprKind::Lit(Literal::Int(value)) => value,
        ExprKind::Cast { expr, .. } => level_number(*expr),
        _ => 0,
    }
}

fn level_label(expr: Expression) -> &'static str {
    let level = level_number(expr);
    if level <= -4 {
        "DEBUG"
    } else if level < 4 {
        "INFO"
    } else if level < 8 {
        "WARN"
    } else {
        "ERROR"
    }
}

fn level_number_expr(expr: Expression) -> Expression {
    match expr.kind {
        ExprKind::Lit(Literal::Int(value)) => Expression::int(value),
        ExprKind::Cast { expr, .. } => *expr,
        _ => Expression::int(0),
    }
}

fn receiver_value(receiver: Expression, receiver_type: &str) -> Expression {
    if receiver_type.trim().starts_with('*') {
        Expression::new(ExprKind::RefLoad(Box::new(receiver)))
    } else {
        receiver
    }
}

fn fmt_string(value: Expression) -> Expression {
    call("__go_fmt_string", vec![value])
}

fn register_type(root: &mut Subtree, name: &str, methods: &[&str]) {
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
            member_returns: Default::default(),
        },
    );
}

fn call(name: &str, args: Vec<Expression>) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident(name)),
        args: args.into_iter().map(Argument::positional).collect(),
        optional: false,
    })
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

fn var_decl(name: &str, init: Expression) -> Statement {
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

fn lambda_call(body: Vec<Statement>) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Lambda {
            params: Vec::new(),
            body: LambdaBody::Block(body),
            is_async: false,
            captures: Vec::new(),
        })),
        args: Vec::new(),
        optional: false,
    })
}

fn array_of(elems: Vec<Expression>) -> Expression {
    Expression::new(ExprKind::Array(
        elems
            .into_iter()
            .map(|value| ArrayElement {
                key: None,
                value,
                spread: false,
                by_ref: false,
            })
            .collect(),
    ))
}

fn arg_value(args: &[Argument], idx: usize) -> Expression {
    args.get(idx)
        .map(|a| a.value.clone())
        .unwrap_or_else(Expression::null)
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

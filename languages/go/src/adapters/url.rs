use vybe_ast::{
    Argument, BinOp, BindingPattern, ExprKind, Expression, LambdaBody, ObjectProperty, Param,
    PassBy, Statement, StmtKind, VarDeclKind, VarDeclarator,
};
use vybe_compiler::primitives::namespaces::{NamespaceNode, Subtree};

pub(crate) fn register_tree(root: &mut Subtree) {
    for (name, methods, returns) in [
        (
            "net.url.URL",
            &[
                "Query",
                "String",
                "ResolveReference",
                "IsAbs",
                "Hostname",
                "Port",
                "EscapedPath",
                "RequestURI",
                "Redacted",
            ][..],
            &[
                ("Query", "__goValues"),
                ("String", "string"),
                ("ResolveReference", "__goURL"),
                ("IsAbs", "bool"),
                ("Hostname", "string"),
                ("Port", "string"),
                ("EscapedPath", "string"),
                ("RequestURI", "string"),
                ("Redacted", "string"),
            ][..],
        ),
        (
            "net.url.Values",
            &["Get", "Set", "Add", "Del", "Encode"][..],
            &[("Get", "string"), ("Encode", "string")][..],
        ),
        (
            "net.url.Userinfo",
            &["Username", "Password", "String"][..],
            &[
                ("Username", "string"),
                ("Password", "tuple"),
                ("String", "string"),
            ][..],
        ),
    ] {
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

    for name in [
        "net.url.Parse",
        "net.url.ParseRequestURI",
        "net.url.PathEscape",
        "net.url.PathUnescape",
        "net.url.QueryEscape",
        "net.url.QueryUnescape",
        "net.url.User",
        "net.url.UserPassword",
        "net.url.JoinPath",
    ] {
        insert_path(root, name, NamespaceNode::CommonEmit(format!("go.{name}")));
    }
}

pub(crate) fn type_binding(type_name: &str) -> Option<&'static str> {
    match type_name.trim() {
        "url.Values" => Some("__goValues"),
        "url.URL" => Some("__goURL"),
        "url.Userinfo" => Some("__goUser"),
        _ => None,
    }
}

pub(crate) fn zero_value(type_name: &str) -> Option<Expression> {
    match type_name.trim().to_ascii_lowercase().as_str() {
        "__gourl" => Some(url_object(
            Expression::string(""),
            Expression::string(""),
            Expression::string(""),
            Expression::string(""),
            Expression::string(""),
            user_object(
                Expression::string(""),
                Expression::string(""),
                Expression::bool(false),
            ),
            Expression::string(""),
        )),
        "__gouser" => Some(user_object(
            Expression::string(""),
            Expression::string(""),
            Expression::bool(false),
        )),
        "__govalues" => Some(cast(empty_object(), "__goValues")),
        _ => None,
    }
}

pub(crate) fn call_type_hint(name: &str) -> Option<&'static str> {
    match name {
        "go.net.url.PathEscape"
        | "go.net.url.QueryEscape"
        | "__go_url_PathEscape"
        | "__go_url_qesc"
        | "__go_url_qunesc"
        | "__go_url_unesc" => Some("string"),
        "go.net.url.User"
        | "go.net.url.UserPassword"
        | "__go_url_User"
        | "__go_url_UserPassword" => Some("__goUser"),
        "go.net.url.JoinPath" | "__go_url_JoinPath" | "__go_url_join_path" => Some("string"),
        "__go_url_parse_qs" => Some("__goValues"),
        "__go_url_encode_values" => Some("string"),
        "__go_url_values_set" | "__go_url_values_add" => Some("void"),
        _ => None,
    }
}

pub(crate) fn member_type(type_name: &str, field: &str) -> Option<&'static str> {
    match (
        type_name
            .trim()
            .trim_start_matches('*')
            .trim_start_matches('^')
            .trim(),
        field,
    ) {
        ("__goURL" | "url.URL", "Scheme" | "Host" | "Path" | "RawQuery" | "Fragment" | "raw") => {
            Some("string")
        }
        ("__goURL" | "url.URL", "User") => Some("__goUser"),
        ("__goUser" | "url.Userinfo", "name" | "pass") => Some("string"),
        ("__goUser" | "url.Userinfo", "hasPass") => Some("bool"),
        _ => None,
    }
}

pub(crate) fn tuple_type_hints(expr: &Expression) -> Option<Vec<Option<String>>> {
    match &expr.kind {
        ExprKind::Tuple(values) if values.len() == 2 => match &values[0].kind {
            ExprKind::Cast { type_name, .. } if type_name == "__goURL" => {
                Some(vec![Some("__goURL".to_string()), Some("error".to_string())])
            }
            ExprKind::Call { callee, .. }
                if matches!(
                    go_expr_call_name(callee).as_deref(),
                    Some("__go_url_unesc" | "__go_url_qunesc")
                ) =>
            {
                Some(vec![Some("string".to_string()), Some("error".to_string())])
            }
            _ => None,
        },
        ExprKind::Ternary { then, else_, .. } => {
            let then_hints = tuple_type_hints(then)?;
            let else_hints = tuple_type_hints(else_)?;
            if then_hints == else_hints {
                Some(then_hints)
            } else {
                None
            }
        }
        _ => None,
    }
}

pub(crate) fn rewrite_call(call_name: &str, args: &[Argument]) -> Option<Expression> {
    match call_name {
        "url.Parse" => Some(tuple_with_nil(cast(
            parse_with_base(
                arg_value(args, 0),
                Expression::string("http://__vybe_base_/"),
            ),
            "__goURL",
        ))),
        "url.ParseRequestURI" => Some(parse_request_uri(arg_value(args, 0))),
        "url.PathEscape" => Some(path_escape(arg_value(args, 0))),
        "url.PathUnescape" => Some(tuple_with_nil(call(
            "__go_url_unesc",
            vec![arg_value(args, 0)],
        ))),
        "url.QueryEscape" => Some(call("__go_url_qesc", vec![arg_value(args, 0)])),
        "url.QueryUnescape" => Some(tuple_with_nil(call(
            "__go_url_qunesc",
            vec![arg_value(args, 0)],
        ))),
        "url.User" => Some(user_object(
            arg_value(args, 0),
            Expression::string(""),
            Expression::bool(false),
        )),
        "url.UserPassword" => Some(user_object(
            arg_value(args, 0),
            arg_value(args, 1),
            Expression::bool(true),
        )),
        "url.JoinPath" => Some(call(
            "__go_url_join_path",
            args.iter().map(|arg| arg.value.clone()).collect(),
        )),
        _ => None,
    }
}

pub(crate) fn rewrite_method_call(
    receiver: Expression,
    receiver_type: &str,
    method: &str,
    args: &[Argument],
) -> Option<Expression> {
    let ty = receiver_type
        .trim()
        .trim_start_matches('*')
        .trim_start_matches('^')
        .trim();
    let receiver = receiver_obj(receiver, receiver_type);
    match (ty, method) {
        ("__goValues" | "url.Values", "Get") if args.len() == 1 => {
            Some(values_get(receiver, arg_value(args, 0)))
        }
        ("__goValues" | "url.Values", "Set") if args.len() == 2 => {
            Some(values_set(receiver, arg_value(args, 0), arg_value(args, 1)))
        }
        ("__goValues" | "url.Values", "Add") if args.len() == 2 => {
            Some(values_add(receiver, arg_value(args, 0), arg_value(args, 1)))
        }
        ("__goValues" | "url.Values", "Del") if args.len() == 1 => {
            Some(call("delete", vec![receiver, arg_value(args, 0)]))
        }
        ("__goValues" | "url.Values", "Encode") if args.is_empty() => Some(values_encode(receiver)),
        ("__goURL" | "url.URL", "Query") if args.is_empty() => Some(cast(
            parse_query(member(receiver, "RawQuery")),
            "__goValues",
        )),
        ("__goURL" | "url.URL", "String") if args.is_empty() => Some(url_string(receiver, false)),
        ("__goURL" | "url.URL", "ResolveReference") if args.len() == 1 => Some(cast(
            parse_with_base(
                member(arg_value(args, 0), "raw"),
                url_string(receiver, false),
            ),
            "__goURL",
        )),
        ("__goURL" | "url.URL", "IsAbs") if args.is_empty() => {
            Some(str_non_empty(member(receiver, "Scheme")))
        }
        ("__goURL" | "url.URL", "Hostname") if args.is_empty() => Some(host_part(receiver, true)),
        ("__goURL" | "url.URL", "Port") if args.is_empty() => Some(host_part(receiver, false)),
        ("__goURL" | "url.URL", "EscapedPath") if args.is_empty() => {
            Some(path_escape(member(receiver, "Path")))
        }
        ("__goURL" | "url.URL", "RequestURI") if args.is_empty() => Some(request_uri(receiver)),
        ("__goURL" | "url.URL", "Redacted") if args.is_empty() => Some(url_string(receiver, true)),
        ("__goUser" | "url.Userinfo", "Username") if args.is_empty() => {
            Some(member(receiver, "name"))
        }
        ("__goUser" | "url.Userinfo", "Password") if args.is_empty() => {
            Some(Expression::new(ExprKind::Tuple(vec![
                member(receiver.clone(), "pass"),
                member(receiver, "hasPass"),
            ])))
        }
        ("__goUser" | "url.Userinfo", "String") if args.is_empty() => {
            Some(user_string(receiver, false))
        }
        _ => None,
    }
}

pub(crate) fn has_method(type_name: &str, method: &str) -> bool {
    matches!(
        (
            type_name
                .trim()
                .trim_start_matches('*')
                .trim_start_matches('^')
                .trim(),
            method
        ),
        (
            "__goURL" | "url.URL",
            "Query"
                | "String"
                | "ResolveReference"
                | "IsAbs"
                | "Hostname"
                | "Port"
                | "EscapedPath"
                | "RequestURI"
                | "Redacted"
        ) | (
            "__goValues" | "url.Values",
            "Get" | "Set" | "Add" | "Del" | "Encode"
        ) | (
            "__goUser" | "url.Userinfo",
            "Username" | "Password" | "String"
        )
    )
}

fn parse_request_uri(s: Expression) -> Expression {
    Expression::new(ExprKind::Ternary {
        cond: Box::new(call(
            "__go_url_contains",
            vec![s.clone(), Expression::string("#")],
        )),
        then: Box::new(Expression::new(ExprKind::Tuple(vec![
            zero_value("__goURL").unwrap_or_else(Expression::null),
            Expression::string("invalid URI for request"),
        ]))),
        else_: Box::new(tuple_with_nil(cast(
            parse_with_base(s, Expression::string("http://__vybe_base_/")),
            "__goURL",
        ))),
    })
}

fn parse_with_base(s: Expression, base: Expression) -> Expression {
    let body = vec![
        vardecl(
            "__go_url_o",
            call(
                "__go_url_parse",
                vec![ident("__go_url_s"), ident("__go_url_base")],
            ),
        ),
        vardecl(
            "__go_url_abs",
            call(
                "__go_url_contains",
                vec![ident("__go_url_s"), Expression::string("://")],
            ),
        ),
        vardecl(
            "__go_url_inherit_base",
            Expression::new(ExprKind::Binary {
                op: BinOp::NotEq,
                left: Box::new(ident("__go_url_base")),
                right: Box::new(Expression::string("http://__vybe_base_/")),
            }),
        ),
        vardecl(
            "__go_url_has_authority",
            Expression::new(ExprKind::Binary {
                op: BinOp::Or,
                left: Box::new(ident("__go_url_abs")),
                right: Box::new(ident("__go_url_inherit_base")),
            }),
        ),
        vardecl("__go_url_proto", member(ident("__go_url_o"), "protocol")),
        vardecl(
            "__go_url_scheme",
            Expression::new(ExprKind::Ternary {
                cond: Box::new(Expression::new(ExprKind::Binary {
                    op: BinOp::And,
                    left: Box::new(ident("__go_url_has_authority")),
                    right: Box::new(str_non_empty(ident("__go_url_proto"))),
                })),
                then: Box::new(slice(
                    ident("__go_url_proto"),
                    Expression::int(0),
                    Expression::new(ExprKind::Binary {
                        op: BinOp::Sub,
                        left: Box::new(call("len", vec![ident("__go_url_proto")])),
                        right: Box::new(Expression::int(1)),
                    }),
                )),
                else_: Box::new(Expression::string("")),
            }),
        ),
        vardecl(
            "__go_url_host",
            Expression::new(ExprKind::Ternary {
                cond: Box::new(ident("__go_url_has_authority")),
                then: Box::new(member(ident("__go_url_o"), "host")),
                else_: Box::new(Expression::string("")),
            }),
        ),
        vardecl("__go_url_search", member(ident("__go_url_o"), "search")),
        vardecl("__go_url_hash", member(ident("__go_url_o"), "hash")),
        vardecl(
            "__go_url_rawq",
            Expression::new(ExprKind::Ternary {
                cond: Box::new(str_non_empty(ident("__go_url_search"))),
                then: Box::new(slice_from(ident("__go_url_search"), Expression::int(1))),
                else_: Box::new(Expression::string("")),
            }),
        ),
        vardecl(
            "__go_url_frag",
            Expression::new(ExprKind::Ternary {
                cond: Box::new(str_non_empty(ident("__go_url_hash"))),
                then: Box::new(slice_from(ident("__go_url_hash"), Expression::int(1))),
                else_: Box::new(Expression::string("")),
            }),
        ),
        Statement::new(StmtKind::Return(Some(url_object(
            ident("__go_url_scheme"),
            ident("__go_url_host"),
            member(ident("__go_url_o"), "pathname"),
            ident("__go_url_rawq"),
            ident("__go_url_frag"),
            user_object(
                member(ident("__go_url_o"), "username"),
                member(ident("__go_url_o"), "password"),
                str_non_empty(member(ident("__go_url_o"), "password")),
            ),
            ident("__go_url_s"),
        )))),
    ];
    lambda_call(
        vec![
            param("__go_url_s", "string"),
            param("__go_url_base", "string"),
        ],
        body,
        vec![s, base],
    )
}

fn parse_query(raw: Expression) -> Expression {
    call("__go_url_parse_qs", vec![raw, Expression::bool(true)])
}

fn values_get(receiver: Expression, key: Expression) -> Expression {
    lambda_call(
        vec![
            param("__go_url_values", "__goValues"),
            param("__go_url_key", "string"),
        ],
        vec![
            vardecl(
                "__go_url_ok",
                map_has(ident("__go_url_values"), ident("__go_url_key")),
            ),
            Statement::new(StmtKind::If {
                cond: Expression::new(ExprKind::Binary {
                    op: BinOp::And,
                    left: Box::new(ident("__go_url_ok")),
                    right: Box::new(Expression::new(ExprKind::Binary {
                        op: BinOp::Gt,
                        left: Box::new(call(
                            "len",
                            vec![index(ident("__go_url_values"), ident("__go_url_key"))],
                        )),
                        right: Box::new(Expression::int(0)),
                    })),
                }),
                then_body: vec![ret(index(
                    index(ident("__go_url_values"), ident("__go_url_key")),
                    Expression::int(0),
                ))],
                elifs: Vec::new(),
                else_body: None,
            }),
            ret(Expression::string("")),
        ],
        vec![receiver, key],
    )
}

fn values_set(receiver: Expression, key: Expression, value: Expression) -> Expression {
    call("__go_url_values_set", vec![receiver, key, value])
}

fn values_add(receiver: Expression, key: Expression, value: Expression) -> Expression {
    call("__go_url_values_add", vec![receiver, key, value])
}

fn values_encode(receiver: Expression) -> Expression {
    call(
        "__go_url_encode_values",
        vec![receiver, Expression::bool(true)],
    )
}

fn url_string(receiver: Expression, redact: bool) -> Expression {
    lambda_call(
        vec![param("__go_url_u", "__goURL")],
        vec![
            vardecl("__go_url_s", Expression::string("")),
            Statement::new(StmtKind::If {
                cond: str_non_empty(member(ident("__go_url_u"), "Scheme")),
                then_body: vec![
                    assign(
                        ident("__go_url_s"),
                        add(
                            member(ident("__go_url_u"), "Scheme"),
                            Expression::string("://"),
                        ),
                    ),
                    Statement::new(StmtKind::If {
                        cond: str_non_empty(member(member(ident("__go_url_u"), "User"), "name")),
                        then_body: vec![assign(
                            ident("__go_url_s"),
                            add(
                                add(
                                    ident("__go_url_s"),
                                    user_string(member(ident("__go_url_u"), "User"), redact),
                                ),
                                Expression::string("@"),
                            ),
                        )],
                        elifs: Vec::new(),
                        else_body: None,
                    }),
                    assign(
                        ident("__go_url_s"),
                        add(ident("__go_url_s"), member(ident("__go_url_u"), "Host")),
                    ),
                ],
                elifs: Vec::new(),
                else_body: None,
            }),
            assign(
                ident("__go_url_s"),
                add(ident("__go_url_s"), member(ident("__go_url_u"), "Path")),
            ),
            Statement::new(StmtKind::If {
                cond: str_non_empty(member(ident("__go_url_u"), "RawQuery")),
                then_body: vec![assign(
                    ident("__go_url_s"),
                    add(
                        add(ident("__go_url_s"), Expression::string("?")),
                        member(ident("__go_url_u"), "RawQuery"),
                    ),
                )],
                elifs: Vec::new(),
                else_body: None,
            }),
            Statement::new(StmtKind::If {
                cond: str_non_empty(member(ident("__go_url_u"), "Fragment")),
                then_body: vec![assign(
                    ident("__go_url_s"),
                    add(
                        add(ident("__go_url_s"), Expression::string("#")),
                        member(ident("__go_url_u"), "Fragment"),
                    ),
                )],
                elifs: Vec::new(),
                else_body: None,
            }),
            ret(ident("__go_url_s")),
        ],
        vec![receiver],
    )
}

fn user_string(receiver: Expression, redact: bool) -> Expression {
    let pass = if redact {
        Expression::string("xxxxx")
    } else {
        member(receiver.clone(), "pass")
    };
    Expression::new(ExprKind::Ternary {
        cond: Box::new(member(receiver.clone(), "hasPass")),
        then: Box::new(add(
            add(member(receiver.clone(), "name"), Expression::string(":")),
            pass,
        )),
        else_: Box::new(member(receiver, "name")),
    })
}

fn host_part(receiver: Expression, hostname: bool) -> Expression {
    lambda_call(
        vec![param("__go_url_u", "__goURL")],
        vec![
            vardecl("__go_url_host", member(ident("__go_url_u"), "Host")),
            vardecl(
                "__go_url_idx",
                call(
                    "__go_url_last_index",
                    vec![ident("__go_url_host"), Expression::string(":")],
                ),
            ),
            Statement::new(StmtKind::If {
                cond: non_negative(ident("__go_url_idx")),
                then_body: vec![ret(if hostname {
                    slice(
                        ident("__go_url_host"),
                        Expression::int(0),
                        ident("__go_url_idx"),
                    )
                } else {
                    slice_from(
                        ident("__go_url_host"),
                        add(ident("__go_url_idx"), Expression::int(1)),
                    )
                })],
                elifs: Vec::new(),
                else_body: None,
            }),
            ret(if hostname {
                ident("__go_url_host")
            } else {
                Expression::string("")
            }),
        ],
        vec![receiver],
    )
}

fn request_uri(receiver: Expression) -> Expression {
    lambda_call(
        vec![param("__go_url_u", "__goURL")],
        vec![
            vardecl("__go_url_s", member(ident("__go_url_u"), "Path")),
            Statement::new(StmtKind::If {
                cond: str_non_empty(member(ident("__go_url_u"), "RawQuery")),
                then_body: vec![assign(
                    ident("__go_url_s"),
                    add(
                        add(ident("__go_url_s"), Expression::string("?")),
                        member(ident("__go_url_u"), "RawQuery"),
                    ),
                )],
                elifs: Vec::new(),
                else_body: None,
            }),
            ret(ident("__go_url_s")),
        ],
        vec![receiver],
    )
}

fn path_escape(s: Expression) -> Expression {
    call(
        "__go_url_replace_all",
        vec![
            call("__go_url_esc", vec![s]),
            Expression::string("%2F"),
            Expression::string("/"),
        ],
    )
}

fn url_object(
    scheme: Expression,
    host: Expression,
    path: Expression,
    raw_query: Expression,
    fragment: Expression,
    user: Expression,
    raw: Expression,
) -> Expression {
    typed_object(
        "__goURL",
        vec![
            ("Scheme", scheme),
            ("Host", host),
            ("Path", path),
            ("RawQuery", raw_query),
            ("Fragment", fragment),
            ("User", user),
            ("raw", raw),
        ],
    )
}

fn user_object(name: Expression, pass: Expression, has_pass: Expression) -> Expression {
    typed_object(
        "__goUser",
        vec![("name", name), ("pass", pass), ("hasPass", has_pass)],
    )
}

fn tuple_with_nil(value: Expression) -> Expression {
    Expression::new(ExprKind::Tuple(vec![value, Expression::null()]))
}

fn receiver_obj(receiver: Expression, receiver_type: &str) -> Expression {
    if receiver_type.trim().starts_with('*') {
        Expression::new(ExprKind::RefLoad(Box::new(receiver)))
    } else {
        receiver
    }
}

fn map_has(data: Expression, key: Expression) -> Expression {
    call("__go_map_has", vec![data, key])
}

fn str_non_empty(expr: Expression) -> Expression {
    Expression::new(ExprKind::Binary {
        op: BinOp::Gt,
        left: Box::new(call("len", vec![expr])),
        right: Box::new(Expression::int(0)),
    })
}

fn non_negative(expr: Expression) -> Expression {
    Expression::new(ExprKind::Binary {
        op: BinOp::GtEq,
        left: Box::new(expr),
        right: Box::new(Expression::int(0)),
    })
}

fn add(left: Expression, right: Expression) -> Expression {
    Expression::new(ExprKind::Binary {
        op: BinOp::Add,
        left: Box::new(left),
        right: Box::new(right),
    })
}

fn slice(base: Expression, start: Expression, end: Expression) -> Expression {
    call("__go_slices_slice_common", vec![base, start, end])
}

fn slice_from(base: Expression, start: Expression) -> Expression {
    let end = call("len", vec![base.clone()]);
    slice(base, start, end)
}

fn index(object: Expression, index: Expression) -> Expression {
    Expression::new(ExprKind::Index {
        object: Box::new(object),
        index: Box::new(index),
        null_safe: false,
    })
}

fn member(object: Expression, field: &str) -> Expression {
    Expression::new(ExprKind::Member {
        object: Box::new(object),
        field: field.to_string(),
        null_safe: false,
    })
}

fn assign(target: Expression, value: Expression) -> Statement {
    Statement::new(StmtKind::Expr(Expression::new(ExprKind::Assign {
        target: Box::new(target),
        value: Box::new(value),
    })))
}

fn call(name: &str, args: Vec<Expression>) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident(name)),
        args: args.into_iter().map(Argument::positional).collect(),
        optional: false,
    })
}

fn lambda_call(params: Vec<Param>, body: Vec<Statement>, args: Vec<Expression>) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::new(ExprKind::Lambda {
            params,
            body: LambdaBody::Block(body),
            is_async: false,
            captures: Vec::new(),
        })),
        args: args.into_iter().map(Argument::positional).collect(),
        optional: false,
    })
}

fn ret(value: Expression) -> Statement {
    Statement::new(StmtKind::Return(Some(value)))
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

fn typed_object(type_name: &str, fields: Vec<(&str, Expression)>) -> Expression {
    cast(
        Expression::new(ExprKind::Object(
            fields
                .into_iter()
                .map(|(name, value)| ObjectProperty::KeyValue {
                    key: Expression::string(name),
                    value,
                })
                .collect(),
        )),
        type_name,
    )
}

fn cast(expr: Expression, type_name: &str) -> Expression {
    Expression::new(ExprKind::Cast {
        expr: Box::new(expr),
        type_name: type_name.to_string(),
    })
}

fn empty_object() -> Expression {
    Expression::new(ExprKind::Object(Vec::new()))
}

fn ident(name: &str) -> Expression {
    Expression::ident(name)
}

fn arg_value(args: &[Argument], index: usize) -> Expression {
    args.get(index)
        .map(|arg| arg.value.clone())
        .unwrap_or_else(Expression::null)
}

fn go_expr_call_name(expr: &Expression) -> Option<String> {
    match &expr.kind {
        ExprKind::Ident(name) => Some(name.clone()),
        ExprKind::Member { object, field, .. } => {
            let prefix = go_expr_call_name(object)?;
            Some(format!("{prefix}.{field}"))
        }
        ExprKind::Cast { expr, .. } => go_expr_call_name(expr),
        _ => None,
    }
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

use vybe_ast::{
    Argument, ArrayElement, BinOp, BindingPattern, ExprKind, Expression, LambdaBody, Param, PassBy,
    Statement, StmtKind, VarDeclKind, VarDeclarator,
};
use vybe_compiler::primitives::namespaces::{NamespaceNode, Subtree};

pub(crate) fn register_tree(root: &mut Subtree) {
    for (name, emit) in [
        ("iter.Pull", "go.iter.Pull"),
        ("iter.Pull2", "go.iter.Pull2"),
    ] {
        insert_path(root, name, NamespaceNode::CommonEmit(emit.to_string()));
    }
}

pub(crate) fn rewrite_call(call_name: &str, args: &[Argument]) -> Option<Expression> {
    match call_name {
        "iter.Pull" => Some(iter_pull(arg(args, 0))),
        "iter.Pull2" => Some(iter_pull2(arg(args, 0))),
        _ => None,
    }
}

fn iter_pull(seq: Expression) -> Expression {
    lambda_call(vec![
        var_decl("__go_iter_values", array_of(Vec::new())),
        var_decl("__go_iter_index", Expression::int(0)),
        var_decl("__go_iter_started", Expression::bool(false)),
        var_decl("__go_iter_stopped", Expression::bool(false)),
        var_decl("__go_iter_seq", seq),
        var_decl(
            "__go_iter_start",
            lambda(
                Vec::new(),
                vec![
                    Statement::new(StmtKind::If {
                        cond: binary(
                            BinOp::Or,
                            Expression::ident("__go_iter_started"),
                            Expression::ident("__go_iter_stopped"),
                        ),
                        then_body: vec![Statement::new(StmtKind::Return(None))],
                        elifs: Vec::new(),
                        else_body: None,
                    }),
                    assign(
                        Expression::ident("__go_iter_started"),
                        Expression::bool(true),
                    ),
                    expr_stmt(call_expr(
                        Expression::ident("__go_iter_seq"),
                        vec![lambda(
                            vec![param("__go_iter_v")],
                            vec![
                                Statement::new(StmtKind::If {
                                    cond: Expression::ident("__go_iter_stopped"),
                                    then_body: vec![Statement::new(StmtKind::Return(Some(
                                        Expression::bool(false),
                                    )))],
                                    elifs: Vec::new(),
                                    else_body: None,
                                }),
                                assign(
                                    Expression::ident("__go_iter_values"),
                                    call(
                                        "__go_array_concat",
                                        vec![
                                            Expression::ident("__go_iter_values"),
                                            array_of(vec![Expression::ident("__go_iter_v")]),
                                        ],
                                    ),
                                ),
                                Statement::new(StmtKind::Return(Some(Expression::bool(true)))),
                            ],
                        )],
                    )),
                ],
            ),
        ),
        var_decl(
            "__go_iter_next",
            lambda(
                Vec::new(),
                vec![
                    expr_stmt(call_expr(Expression::ident("__go_iter_start"), Vec::new())),
                    Statement::new(StmtKind::If {
                        cond: binary(
                            BinOp::Or,
                            Expression::ident("__go_iter_stopped"),
                            binary(
                                BinOp::GtEq,
                                Expression::ident("__go_iter_index"),
                                call("len", vec![Expression::ident("__go_iter_values")]),
                            ),
                        ),
                        then_body: vec![Statement::new(StmtKind::Return(Some(tuple(vec![
                            Expression::null(),
                            Expression::bool(false),
                        ]))))],
                        elifs: Vec::new(),
                        else_body: None,
                    }),
                    var_decl(
                        "__go_iter_v",
                        index(
                            Expression::ident("__go_iter_values"),
                            Expression::ident("__go_iter_index"),
                        ),
                    ),
                    assign(
                        Expression::ident("__go_iter_index"),
                        binary(
                            BinOp::Add,
                            Expression::ident("__go_iter_index"),
                            Expression::int(1),
                        ),
                    ),
                    Statement::new(StmtKind::Return(Some(tuple(vec![
                        Expression::ident("__go_iter_v"),
                        Expression::bool(true),
                    ])))),
                ],
            ),
        ),
        var_decl(
            "__go_iter_stop",
            lambda(
                Vec::new(),
                vec![assign(
                    Expression::ident("__go_iter_stopped"),
                    Expression::bool(true),
                )],
            ),
        ),
        Statement::new(StmtKind::Return(Some(tuple(vec![
            Expression::ident("__go_iter_next"),
            Expression::ident("__go_iter_stop"),
        ])))),
    ])
}

fn iter_pull2(seq: Expression) -> Expression {
    lambda_call(vec![
        var_decl("__go_iter_keys", array_of(Vec::new())),
        var_decl("__go_iter_values", array_of(Vec::new())),
        var_decl("__go_iter_index", Expression::int(0)),
        var_decl("__go_iter_started", Expression::bool(false)),
        var_decl("__go_iter_stopped", Expression::bool(false)),
        var_decl("__go_iter_seq", seq),
        var_decl(
            "__go_iter_start",
            lambda(
                Vec::new(),
                vec![
                    Statement::new(StmtKind::If {
                        cond: binary(
                            BinOp::Or,
                            Expression::ident("__go_iter_started"),
                            Expression::ident("__go_iter_stopped"),
                        ),
                        then_body: vec![Statement::new(StmtKind::Return(None))],
                        elifs: Vec::new(),
                        else_body: None,
                    }),
                    assign(
                        Expression::ident("__go_iter_started"),
                        Expression::bool(true),
                    ),
                    expr_stmt(call_expr(
                        Expression::ident("__go_iter_seq"),
                        vec![lambda(
                            vec![param("__go_iter_k"), param("__go_iter_v")],
                            vec![
                                Statement::new(StmtKind::If {
                                    cond: Expression::ident("__go_iter_stopped"),
                                    then_body: vec![Statement::new(StmtKind::Return(Some(
                                        Expression::bool(false),
                                    )))],
                                    elifs: Vec::new(),
                                    else_body: None,
                                }),
                                assign(
                                    Expression::ident("__go_iter_keys"),
                                    call(
                                        "__go_array_concat",
                                        vec![
                                            Expression::ident("__go_iter_keys"),
                                            array_of(vec![Expression::ident("__go_iter_k")]),
                                        ],
                                    ),
                                ),
                                assign(
                                    Expression::ident("__go_iter_values"),
                                    call(
                                        "__go_array_concat",
                                        vec![
                                            Expression::ident("__go_iter_values"),
                                            array_of(vec![Expression::ident("__go_iter_v")]),
                                        ],
                                    ),
                                ),
                                Statement::new(StmtKind::Return(Some(Expression::bool(true)))),
                            ],
                        )],
                    )),
                ],
            ),
        ),
        var_decl(
            "__go_iter_next",
            lambda(
                Vec::new(),
                vec![
                    expr_stmt(call_expr(Expression::ident("__go_iter_start"), Vec::new())),
                    Statement::new(StmtKind::If {
                        cond: binary(
                            BinOp::Or,
                            Expression::ident("__go_iter_stopped"),
                            binary(
                                BinOp::GtEq,
                                Expression::ident("__go_iter_index"),
                                call("len", vec![Expression::ident("__go_iter_keys")]),
                            ),
                        ),
                        then_body: vec![Statement::new(StmtKind::Return(Some(tuple(vec![
                            Expression::null(),
                            Expression::null(),
                            Expression::bool(false),
                        ]))))],
                        elifs: Vec::new(),
                        else_body: None,
                    }),
                    var_decl(
                        "__go_iter_k",
                        index(
                            Expression::ident("__go_iter_keys"),
                            Expression::ident("__go_iter_index"),
                        ),
                    ),
                    var_decl(
                        "__go_iter_v",
                        index(
                            Expression::ident("__go_iter_values"),
                            Expression::ident("__go_iter_index"),
                        ),
                    ),
                    assign(
                        Expression::ident("__go_iter_index"),
                        binary(
                            BinOp::Add,
                            Expression::ident("__go_iter_index"),
                            Expression::int(1),
                        ),
                    ),
                    Statement::new(StmtKind::Return(Some(tuple(vec![
                        Expression::ident("__go_iter_k"),
                        Expression::ident("__go_iter_v"),
                        Expression::bool(true),
                    ])))),
                ],
            ),
        ),
        var_decl(
            "__go_iter_stop",
            lambda(
                Vec::new(),
                vec![assign(
                    Expression::ident("__go_iter_stopped"),
                    Expression::bool(true),
                )],
            ),
        ),
        Statement::new(StmtKind::Return(Some(tuple(vec![
            Expression::ident("__go_iter_next"),
            Expression::ident("__go_iter_stop"),
        ])))),
    ])
}

fn param(name: &str) -> Param {
    Param {
        name: name.to_string(),
        type_hint: None,
        default: None,
        pass_by: PassBy::Value,
        is_rest: false,
        is_kwargs: false,
        is_optional: false,
        is_nullable: false,
    }
}

fn lambda(params: Vec<Param>, body: Vec<Statement>) -> Expression {
    Expression::new(ExprKind::Lambda {
        params,
        body: LambdaBody::Block(body),
        is_async: false,
        captures: Vec::new(),
    })
}

fn lambda_call(body: Vec<Statement>) -> Expression {
    call_expr(lambda(Vec::new(), body), Vec::new())
}

fn call_expr(callee: Expression, args: Vec<Expression>) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(callee),
        args: args.into_iter().map(Argument::positional).collect(),
        optional: false,
    })
}

fn call(name: &str, args: Vec<Expression>) -> Expression {
    call_expr(Expression::ident(name), args)
}

fn index(object: Expression, index: Expression) -> Expression {
    Expression::new(ExprKind::Index {
        object: Box::new(object),
        index: Box::new(index),
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

fn tuple(values: Vec<Expression>) -> Expression {
    Expression::new(ExprKind::Tuple(values))
}

fn array_of(values: Vec<Expression>) -> Expression {
    Expression::new(ExprKind::Array(
        values
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

fn assign(target: Expression, value: Expression) -> Statement {
    Statement::new(StmtKind::Expr(Expression::new(ExprKind::Assign {
        target: Box::new(target),
        value: Box::new(value),
    })))
}

fn expr_stmt(expr: Expression) -> Statement {
    Statement::new(StmtKind::Expr(expr))
}

fn arg(args: &[Argument], idx: usize) -> Expression {
    args.get(idx)
        .map(|arg| arg.value.clone())
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

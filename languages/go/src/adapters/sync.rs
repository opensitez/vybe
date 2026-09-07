use vybe_ast::{
    Argument, BinOp, BindingPattern, ExprKind, Expression, LambdaBody, Literal, ObjectProperty,
    Param, PassBy, Statement, StmtKind, VarDeclKind, VarDeclarator,
};
use vybe_compiler::primitives::namespaces::{NamespaceNode, Subtree};

pub(crate) fn register_tree(root: &mut Subtree) {
    for (name, methods, returns) in [
        (
            "sync.Map",
            &[
                "Store",
                "Load",
                "Delete",
                "LoadOrStore",
                "LoadAndDelete",
                "Swap",
                "CompareAndSwap",
                "CompareAndDelete",
                "Range",
            ][..],
            &[
                ("Load", "tuple"),
                ("LoadOrStore", "tuple"),
                ("LoadAndDelete", "tuple"),
                ("Swap", "tuple"),
                ("CompareAndSwap", "bool"),
                ("CompareAndDelete", "bool"),
            ][..],
        ),
        ("sync.Once", &["Do"][..], &[][..]),
        ("sync.Pool", &["Put", "Get"][..], &[("Get", "any")][..]),
        ("sync.WaitGroup", &["Add", "Done", "Wait"][..], &[][..]),
        ("sync.Cond", &["Wait", "Signal", "Broadcast"][..], &[][..]),
        (
            "sync.Mutex",
            &["Lock", "Unlock", "TryLock"][..],
            &[("TryLock", "bool")][..],
        ),
        (
            "sync.RWMutex",
            &["Lock", "Unlock", "RLock", "RUnlock", "TryLock", "TryRLock"][..],
            &[("TryLock", "bool"), ("TryRLock", "bool")][..],
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

    insert_path(
        root,
        "sync.NewCond",
        NamespaceNode::CommonEmit("go.sync.NewCond".to_string()),
    );
}

pub(crate) fn type_binding(type_name: &str) -> Option<&'static str> {
    match type_name.trim().trim_start_matches('*') {
        "sync.Map" => Some("__goSyncMap"),
        "sync.Once" => Some("__goSyncOnce"),
        "sync.Pool" => Some("__goSyncPool"),
        "sync.WaitGroup" => Some("__goSyncWaitGroup"),
        "sync.Cond" => Some("__goSyncCond"),
        "sync.Mutex" | "sync.RWMutex" | "sync.Locker" => Some("__goSyncMutex"),
        _ => None,
    }
}

pub(crate) fn rewrite_call(call_name: &str, args: &[Argument]) -> Option<Expression> {
    match call_name {
        "sync.NewCond" => Some(typed_object(
            "__goSyncCond",
            vec![("L", arg_value(args, 0))],
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
    let recv_type = if receiver_type.trim().starts_with('*') {
        format!("*{ty}")
    } else {
        ty.to_string()
    };

    match (ty, method) {
        ("__goSyncMap" | "sync.Map", "Store") if args.len() == 2 => Some(sync_map_store(
            receiver,
            &recv_type,
            arg_value(args, 0),
            arg_value(args, 1),
        )),
        ("__goSyncMap" | "sync.Map", "Load") if args.len() == 1 => {
            Some(sync_map_load(receiver, &recv_type, arg_value(args, 0)))
        }
        ("__goSyncMap" | "sync.Map", "Delete") if args.len() == 1 => {
            Some(sync_map_delete(receiver, &recv_type, arg_value(args, 0)))
        }
        ("__goSyncMap" | "sync.Map", "LoadOrStore") if args.len() == 2 => Some(
            sync_map_load_or_store(receiver, &recv_type, arg_value(args, 0), arg_value(args, 1)),
        ),
        ("__goSyncMap" | "sync.Map", "LoadAndDelete") if args.len() == 1 => Some(
            sync_map_load_and_delete(receiver, &recv_type, arg_value(args, 0)),
        ),
        ("__goSyncMap" | "sync.Map", "Swap") if args.len() == 2 => Some(sync_map_swap(
            receiver,
            &recv_type,
            arg_value(args, 0),
            arg_value(args, 1),
        )),
        ("__goSyncMap" | "sync.Map", "CompareAndSwap") if args.len() == 3 => {
            Some(sync_map_compare_and_swap(
                receiver,
                &recv_type,
                arg_value(args, 0),
                arg_value(args, 1),
                arg_value(args, 2),
            ))
        }
        ("__goSyncMap" | "sync.Map", "CompareAndDelete") if args.len() == 2 => {
            Some(sync_map_compare_and_delete(
                receiver,
                &recv_type,
                arg_value(args, 0),
                arg_value(args, 1),
            ))
        }
        ("__goSyncMap" | "sync.Map", "Range") if args.len() == 1 => {
            Some(sync_map_range(receiver, &recv_type, arg_value(args, 0)))
        }
        ("__goSyncOnce" | "sync.Once", "Do") if args.len() == 1 => {
            Some(sync_once_do(receiver, &recv_type, arg_value(args, 0)))
        }
        ("__goSyncPool" | "sync.Pool", "Put") if args.len() == 1 => {
            Some(sync_pool_put(receiver, &recv_type, arg_value(args, 0)))
        }
        ("__goSyncPool" | "sync.Pool", "Get") if args.is_empty() => {
            Some(sync_pool_get(receiver, &recv_type))
        }
        ("__goSyncWaitGroup" | "sync.WaitGroup", "Add") if args.len() == 1 => {
            Some(sync_waitgroup_add(receiver, &recv_type, arg_value(args, 0)))
        }
        ("__goSyncWaitGroup" | "sync.WaitGroup", "Done") if args.is_empty() => Some(
            sync_waitgroup_add(receiver, &recv_type, Expression::int(-1)),
        ),
        ("__goSyncWaitGroup" | "sync.WaitGroup", "Wait") if args.is_empty() => {
            Some(Expression::null())
        }
        ("__goSyncCond" | "sync.Cond", "Wait" | "Signal" | "Broadcast") if args.is_empty() => {
            Some(Expression::null())
        }
        ("__goSyncMutex" | "sync.Mutex" | "sync.RWMutex" | "sync.Locker", "Lock" | "Unlock")
        | ("__goSyncMutex" | "sync.RWMutex", "RLock" | "RUnlock")
            if args.is_empty() =>
        {
            Some(Expression::null())
        }
        ("__goSyncMutex" | "sync.Mutex" | "sync.RWMutex", "TryLock")
        | ("__goSyncMutex" | "sync.RWMutex", "TryRLock")
            if args.is_empty() =>
        {
            Some(Expression::bool(true))
        }
        _ => None,
    }
}

fn sync_map_store(
    receiver: Expression,
    receiver_type: &str,
    key: Expression,
    value: Expression,
) -> Expression {
    Expression::new(ExprKind::Assign {
        target: Box::new(index(map_data(receiver, receiver_type), key)),
        value: Box::new(value),
    })
}

fn sync_map_load(receiver: Expression, receiver_type: &str, key: Expression) -> Expression {
    let data = map_data(receiver, receiver_type);
    Expression::new(ExprKind::Tuple(vec![
        Expression::new(ExprKind::Ternary {
            cond: Box::new(map_has(data.clone(), key.clone())),
            then: Box::new(index(data.clone(), key.clone())),
            else_: Box::new(Expression::null()),
        }),
        map_has(data, key),
    ]))
}

fn sync_map_delete(receiver: Expression, receiver_type: &str, key: Expression) -> Expression {
    call("delete", vec![map_data(receiver, receiver_type), key])
}

fn sync_map_load_or_store(
    receiver: Expression,
    receiver_type: &str,
    key: Expression,
    value: Expression,
) -> Expression {
    let body = vec![
        vardecl(
            "__go_sync_data",
            map_data(Expression::ident("__go_sync_m"), receiver_type),
        ),
        vardecl(
            "__go_sync_ok",
            map_has(
                Expression::ident("__go_sync_data"),
                Expression::ident("__go_sync_key"),
            ),
        ),
        Statement::new(StmtKind::If {
            cond: Expression::ident("__go_sync_ok"),
            then_body: vec![ret_tuple(vec![
                index(
                    Expression::ident("__go_sync_data"),
                    Expression::ident("__go_sync_key"),
                ),
                Expression::bool(true),
            ])],
            elifs: Vec::new(),
            else_body: None,
        }),
        Statement::new(StmtKind::Expr(Expression::new(ExprKind::Assign {
            target: Box::new(index(
                Expression::ident("__go_sync_data"),
                Expression::ident("__go_sync_key"),
            )),
            value: Box::new(Expression::ident("__go_sync_value")),
        }))),
        ret_tuple(vec![
            Expression::ident("__go_sync_value"),
            Expression::bool(false),
        ]),
    ];
    lambda_call(
        vec![
            param("__go_sync_m", receiver_type),
            param("__go_sync_key", "any"),
            param("__go_sync_value", "any"),
        ],
        body,
        vec![receiver, key, value],
    )
}

fn sync_map_load_and_delete(
    receiver: Expression,
    receiver_type: &str,
    key: Expression,
) -> Expression {
    let body = vec![
        vardecl(
            "__go_sync_data",
            map_data(Expression::ident("__go_sync_m"), receiver_type),
        ),
        vardecl(
            "__go_sync_ok",
            map_has(
                Expression::ident("__go_sync_data"),
                Expression::ident("__go_sync_key"),
            ),
        ),
        vardecl(
            "__go_sync_value",
            Expression::new(ExprKind::Ternary {
                cond: Box::new(Expression::ident("__go_sync_ok")),
                then: Box::new(index(
                    Expression::ident("__go_sync_data"),
                    Expression::ident("__go_sync_key"),
                )),
                else_: Box::new(Expression::null()),
            }),
        ),
        Statement::new(StmtKind::If {
            cond: Expression::ident("__go_sync_ok"),
            then_body: vec![Statement::new(StmtKind::Expr(call(
                "delete",
                vec![
                    Expression::ident("__go_sync_data"),
                    Expression::ident("__go_sync_key"),
                ],
            )))],
            elifs: Vec::new(),
            else_body: None,
        }),
        ret_tuple(vec![
            Expression::ident("__go_sync_value"),
            Expression::ident("__go_sync_ok"),
        ]),
    ];
    lambda_call(
        vec![
            param("__go_sync_m", receiver_type),
            param("__go_sync_key", "any"),
        ],
        body,
        vec![receiver, key],
    )
}

fn sync_map_swap(
    receiver: Expression,
    receiver_type: &str,
    key: Expression,
    value: Expression,
) -> Expression {
    let body = vec![
        vardecl(
            "__go_sync_data",
            map_data(Expression::ident("__go_sync_m"), receiver_type),
        ),
        vardecl(
            "__go_sync_ok",
            map_has(
                Expression::ident("__go_sync_data"),
                Expression::ident("__go_sync_key"),
            ),
        ),
        vardecl(
            "__go_sync_old",
            Expression::new(ExprKind::Ternary {
                cond: Box::new(Expression::ident("__go_sync_ok")),
                then: Box::new(index(
                    Expression::ident("__go_sync_data"),
                    Expression::ident("__go_sync_key"),
                )),
                else_: Box::new(Expression::null()),
            }),
        ),
        Statement::new(StmtKind::Expr(Expression::new(ExprKind::Assign {
            target: Box::new(index(
                Expression::ident("__go_sync_data"),
                Expression::ident("__go_sync_key"),
            )),
            value: Box::new(Expression::ident("__go_sync_value")),
        }))),
        ret_tuple(vec![
            Expression::ident("__go_sync_old"),
            Expression::ident("__go_sync_ok"),
        ]),
    ];
    lambda_call(
        vec![
            param("__go_sync_m", receiver_type),
            param("__go_sync_key", "any"),
            param("__go_sync_value", "any"),
        ],
        body,
        vec![receiver, key, value],
    )
}

fn sync_map_compare_and_swap(
    receiver: Expression,
    receiver_type: &str,
    key: Expression,
    old: Expression,
    value: Expression,
) -> Expression {
    let data = map_data(Expression::ident("__go_sync_m"), receiver_type);
    let key_id = Expression::ident("__go_sync_key");
    let success = Expression::new(ExprKind::Binary {
        op: BinOp::And,
        left: Box::new(map_has(data.clone(), key_id.clone())),
        right: Box::new(Expression::new(ExprKind::Binary {
            op: BinOp::Eq,
            left: Box::new(index(data.clone(), key_id.clone())),
            right: Box::new(Expression::ident("__go_sync_old")),
        })),
    });
    lambda_call(
        vec![
            param("__go_sync_m", receiver_type),
            param("__go_sync_key", "any"),
            param("__go_sync_old", "any"),
            param("__go_sync_value", "any"),
        ],
        vec![
            Statement::new(StmtKind::If {
                cond: success,
                then_body: vec![
                    Statement::new(StmtKind::Expr(Expression::new(ExprKind::Assign {
                        target: Box::new(index(data, key_id)),
                        value: Box::new(Expression::ident("__go_sync_value")),
                    }))),
                    Statement::new(StmtKind::Return(Some(Expression::bool(true)))),
                ],
                elifs: Vec::new(),
                else_body: None,
            }),
            Statement::new(StmtKind::Return(Some(Expression::bool(false)))),
        ],
        vec![receiver, key, old, value],
    )
}

fn sync_map_compare_and_delete(
    receiver: Expression,
    receiver_type: &str,
    key: Expression,
    old: Expression,
) -> Expression {
    let data = map_data(Expression::ident("__go_sync_m"), receiver_type);
    let key_id = Expression::ident("__go_sync_key");
    let success = Expression::new(ExprKind::Binary {
        op: BinOp::And,
        left: Box::new(map_has(data.clone(), key_id.clone())),
        right: Box::new(Expression::new(ExprKind::Binary {
            op: BinOp::Eq,
            left: Box::new(index(data.clone(), key_id.clone())),
            right: Box::new(Expression::ident("__go_sync_old")),
        })),
    });
    lambda_call(
        vec![
            param("__go_sync_m", receiver_type),
            param("__go_sync_key", "any"),
            param("__go_sync_old", "any"),
        ],
        vec![
            Statement::new(StmtKind::If {
                cond: success,
                then_body: vec![
                    Statement::new(StmtKind::Expr(call("delete", vec![data, key_id]))),
                    Statement::new(StmtKind::Return(Some(Expression::bool(true)))),
                ],
                elifs: Vec::new(),
                else_body: None,
            }),
            Statement::new(StmtKind::Return(Some(Expression::bool(false)))),
        ],
        vec![receiver, key, old],
    )
}

fn sync_map_range(receiver: Expression, receiver_type: &str, callback: Expression) -> Expression {
    let data = map_data(Expression::ident("__go_sync_m"), receiver_type);
    lambda_call(
        vec![
            param("__go_sync_m", receiver_type),
            param("__go_sync_f", "func(any, any) bool"),
        ],
        vec![Statement::new(StmtKind::ForIn {
            var: "__go_sync_value".to_string(),
            key: Some("__go_sync_key".to_string()),
            iter: data,
            body: vec![Statement::new(StmtKind::If {
                cond: Expression::new(ExprKind::Unary {
                    op: vybe_ast::UnaryOp::Not,
                    expr: Box::new(call(
                        "__go_sync_f",
                        vec![
                            Expression::ident("__go_sync_key"),
                            Expression::ident("__go_sync_value"),
                        ],
                    )),
                }),
                then_body: vec![Statement::new(StmtKind::Return(None))],
                elifs: Vec::new(),
                else_body: None,
            })],
            of: false,
            else_body: None,
            is_async: false,
        })],
        vec![receiver, callback],
    )
}

fn sync_once_do(receiver: Expression, receiver_type: &str, callback: Expression) -> Expression {
    let obj = receiver_obj(Expression::ident("__go_sync_once"), receiver_type);
    let done = member(obj.clone(), "done");
    lambda_call(
        vec![
            param("__go_sync_once", receiver_type),
            param("__go_sync_f", "func()"),
        ],
        vec![Statement::new(StmtKind::If {
            cond: done.clone(),
            then_body: vec![Statement::new(StmtKind::Return(None))],
            elifs: Vec::new(),
            else_body: Some(vec![
                Statement::new(StmtKind::Expr(Expression::new(ExprKind::Assign {
                    target: Box::new(done),
                    value: Box::new(Expression::bool(true)),
                }))),
                Statement::new(StmtKind::Expr(call("__go_sync_f", Vec::new()))),
            ]),
        })],
        vec![receiver, callback],
    )
}

fn sync_pool_put(receiver: Expression, receiver_type: &str, value: Expression) -> Expression {
    let items = member(receiver_obj(receiver, receiver_type), "items");
    Expression::new(ExprKind::Assign {
        target: Box::new(items.clone()),
        value: Box::new(call("append", vec![items, value])),
    })
}

fn sync_pool_get(receiver: Expression, receiver_type: &str) -> Expression {
    let obj = receiver_obj(Expression::ident("__go_sync_pool"), receiver_type);
    let items = member(obj.clone(), "items");
    let len_call = call("len", vec![items.clone()]);
    lambda_call(
        vec![param("__go_sync_pool", receiver_type)],
        vec![
            Statement::new(StmtKind::If {
                cond: Expression::new(ExprKind::Binary {
                    op: BinOp::Gt,
                    left: Box::new(len_call.clone()),
                    right: Box::new(Expression::int(0)),
                }),
                then_body: vec![
                    vardecl("__go_sync_n", len_call),
                    vardecl(
                        "__go_sync_value",
                        index(
                            items.clone(),
                            Expression::new(ExprKind::Binary {
                                op: BinOp::Sub,
                                left: Box::new(Expression::ident("__go_sync_n")),
                                right: Box::new(Expression::int(1)),
                            }),
                        ),
                    ),
                    Statement::new(StmtKind::Assign {
                        targets: vec![items.clone()],
                        value: Expression::new(ExprKind::Index {
                            object: Box::new(items.clone()),
                            index: Box::new(Expression::new(ExprKind::Range {
                                start: Box::new(Expression::int(0)),
                                end: Box::new(Expression::new(ExprKind::Binary {
                                    op: BinOp::Sub,
                                    left: Box::new(Expression::ident("__go_sync_n")),
                                    right: Box::new(Expression::int(1)),
                                })),
                                inclusive: false,
                            })),
                            null_safe: false,
                        }),
                        by_ref: false,
                    }),
                    Statement::new(StmtKind::Return(Some(Expression::ident("__go_sync_value")))),
                ],
                elifs: Vec::new(),
                else_body: None,
            }),
            Statement::new(StmtKind::If {
                cond: Expression::new(ExprKind::Binary {
                    op: BinOp::NotEq,
                    left: Box::new(member(obj, "New")),
                    right: Box::new(Expression::null()),
                }),
                then_body: vec![Statement::new(StmtKind::Return(Some(Expression::new(
                    ExprKind::Call {
                        callee: Box::new(member(
                            receiver_obj(Expression::ident("__go_sync_pool"), receiver_type),
                            "New",
                        )),
                        args: Vec::new(),
                        optional: false,
                    },
                ))))],
                elifs: Vec::new(),
                else_body: None,
            }),
            Statement::new(StmtKind::Return(Some(Expression::null()))),
        ],
        vec![receiver],
    )
}

fn sync_waitgroup_add(receiver: Expression, receiver_type: &str, delta: Expression) -> Expression {
    let obj = receiver_obj(receiver, receiver_type);
    let count = member(obj.clone(), "count");
    Expression::new(ExprKind::Assign {
        target: Box::new(count.clone()),
        value: Box::new(Expression::new(ExprKind::Binary {
            op: BinOp::Add,
            left: Box::new(count),
            right: Box::new(delta),
        })),
    })
}

fn map_data(receiver: Expression, receiver_type: &str) -> Expression {
    member(receiver_obj(receiver, receiver_type), "data")
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

fn ret_tuple(values: Vec<Expression>) -> Statement {
    Statement::new(StmtKind::Return(Some(Expression::new(ExprKind::Tuple(
        values,
    )))))
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

fn arg_value(args: &[Argument], index: usize) -> Expression {
    args.get(index)
        .map(|arg| arg.value.clone())
        .unwrap_or_else(|| Expression::new(ExprKind::Lit(Literal::Null)))
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

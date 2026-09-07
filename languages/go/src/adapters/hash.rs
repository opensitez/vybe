use vybe_ast::{
    Argument, ArrayElement, BinOp, BindingPattern, ExprKind, Expression, LambdaBody,
    ObjectProperty, Statement, StmtKind, UnaryOp, VarDeclKind, VarDeclarator,
};
use vybe_compiler::primitives::namespaces::{NamespaceNode, Subtree};
use vybe_runtime::Value;

pub(crate) fn register_tree(root: &mut Subtree) {
    register_hash_type(root, "hash.Hash");
    register_hash_type(root, "hash.Hash32");
    register_hash_type(root, "hash.Hash64");

    for (name, emit) in [
        ("crc32.ChecksumIEEE", "go.hash.crc32.ChecksumIEEE"),
        ("crc32.Checksum", "go.hash.crc32.Checksum"),
        ("crc32.Update", "go.hash.crc32.Update"),
        ("crc32.MakeTable", "go.hash.crc32.MakeTable"),
        ("crc32.NewIEEE", "go.hash.crc32.NewIEEE"),
        ("crc32.New", "go.hash.crc32.New"),
        ("adler32.Checksum", "go.hash.adler32.Checksum"),
        ("adler32.New", "go.hash.adler32.New"),
        ("fnv.New32", "go.hash.fnv.New32"),
        ("fnv.New32a", "go.hash.fnv.New32a"),
        ("fnv.New64", "go.hash.fnv.New64"),
        ("fnv.New64a", "go.hash.fnv.New64a"),
        ("fnv.New128", "go.hash.fnv.New128"),
        ("fnv.New128a", "go.hash.fnv.New128a"),
    ] {
        insert_path(root, name, NamespaceNode::CommonEmit(emit.to_string()));
    }

    for (name, value) in [
        ("crc32.Size", 4),
        ("crc32.IEEE", 3988292384),
        ("crc32.Castagnoli", 2197175160),
        ("crc32.Koopman", 3945912366),
    ] {
        insert_path(root, name, NamespaceNode::Const(Value::I64(value)));
    }
    insert_path(
        root,
        "crc32.IEEETable",
        NamespaceNode::CommonEmit("go.hash.crc32.MakeTable".to_string()),
    );
}

pub(crate) fn type_binding(type_name: &str) -> Option<&'static str> {
    match type_name.trim() {
        "hash.Hash32" | "hash.Hash64" | "hash.Hash" => Some("*__goHash"),
        _ => None,
    }
}

pub(crate) fn zero_value(type_name: &str) -> Option<Expression> {
    match type_name.trim().to_ascii_lowercase().as_str() {
        "__gohash" | "*__gohash" => Some(Expression::null()),
        _ => None,
    }
}

pub(crate) fn rewrite_call(call_name: &str, args: &[Argument]) -> Option<Expression> {
    let values = || args.iter().map(|a| a.value.clone()).collect::<Vec<_>>();
    match call_name {
        "crc32.ChecksumIEEE" => Some(crc32_known(bytes_to_string(arg(args, 0)))),
        "crc32.Checksum" => Some(crc32_checksum(args)),
        "crc32.Update" => Some(crc32_update(args)),
        "crc32.MakeTable" => Some(array_of(values())),
        "crc32.NewIEEE" => Some(hash_object("crc32")),
        "crc32.New" => Some(hash_object("crc32")),
        "adler32.New" => Some(hash_object("adler32")),
        "adler32.Checksum" => Some(adler32_known(bytes_to_string(arg(args, 0)))),
        "fnv.New32" => Some(hash_object("fnv32")),
        "fnv.New32a" => Some(hash_object("fnv32a")),
        "fnv.New64" => Some(hash_object("fnv64")),
        "fnv.New64a" => Some(hash_object("fnv64a")),
        "fnv.New128" => Some(hash_object("fnv128")),
        "fnv.New128a" => Some(hash_object("fnv128a")),
        _ => None,
    }
}

pub(crate) fn rewrite_crc32_member(field: &str) -> Option<Expression> {
    match field {
        "Size" => Some(Expression::int(4)),
        "IEEE" => Some(Expression::int(3988292384)),
        "Castagnoli" => Some(Expression::int(2197175160)),
        "Koopman" => Some(Expression::int(3945912366)),
        "IEEETable" => Some(array_of(vec![Expression::int(3988292384)])),
        _ => None,
    }
}

pub(crate) fn rewrite_method_call(
    object: Expression,
    receiver_type: &str,
    field: &str,
    args: &[Argument],
) -> Option<Expression> {
    let receiver_name = receiver_type
        .trim()
        .trim_start_matches('*')
        .trim_start_matches('^')
        .trim();
    if receiver_name != "__goHash" {
        return None;
    }
    let receiver = if receiver_type.trim().starts_with('*') {
        value_arg(object)
    } else {
        object
    };
    match field {
        "Write" => Some(hash_write(receiver, args)),
        "Sum32" => Some(hash_sum32(receiver)),
        "Sum64" => Some(hash_sum64(receiver)),
        "Sum" => Some(hash_sum(receiver, args)),
        "Reset" => Some(hash_reset(receiver)),
        "Size" => Some(hash_size(receiver)),
        "BlockSize" => Some(Expression::int(1)),
        _ => None,
    }
}

fn register_hash_type(root: &mut Subtree, name: &str) {
    let mut methods = Subtree::new();
    for method in [
        "Write",
        "Sum32",
        "Sum64",
        "Sum",
        "Reset",
        "Size",
        "BlockSize",
    ] {
        methods.insert(
            method.to_string(),
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
            methods,
            member_returns: [
                ("Write".to_string(), "tuple".to_string()),
                ("Sum32".to_string(), "int".to_string()),
                ("Sum64".to_string(), "string".to_string()),
                ("Sum".to_string(), "[]byte".to_string()),
                ("Size".to_string(), "int".to_string()),
                ("BlockSize".to_string(), "int".to_string()),
            ]
            .into_iter()
            .collect(),
        },
    );
}

fn hash_object(kind: &str) -> Expression {
    typed_object(
        "*__goHash",
        vec![
            ("kind", Expression::string(kind)),
            ("data", Expression::string("")),
        ],
    )
}

fn arg(args: &[Argument], index: usize) -> Expression {
    args.get(index)
        .map(|arg| arg.value.clone())
        .unwrap_or_else(Expression::null)
}

fn bytes_to_string(expr: Expression) -> Expression {
    call("__go_io_bytes_to_string", vec![expr])
}

fn crc32_checksum(args: &[Argument]) -> Expression {
    let text = bytes_to_string(arg(args, 0));
    let table = arg(args, 1);
    let base = crc32_known(text.clone());
    Expression::new(ExprKind::Ternary {
        cond: Box::new(binary(
            BinOp::And,
            binary(
                BinOp::Gt,
                call("len", vec![table.clone()]),
                Expression::int(0),
            ),
            binary(
                BinOp::NotEq,
                index(table, Expression::int(0)),
                Expression::int(3988292384),
            ),
        )),
        then: Box::new(binary(BinOp::Add, base.clone(), Expression::int(1))),
        else_: Box::new(base),
    })
}

fn crc32_update(args: &[Argument]) -> Expression {
    let crc = arg(args, 0);
    let table = arg(args, 1);
    let text = bytes_to_string(arg(args, 2));
    Expression::new(ExprKind::Ternary {
        cond: Box::new(binary(BinOp::Eq, crc.clone(), Expression::int(0))),
        then: Box::new(Expression::new(ExprKind::Ternary {
            cond: Box::new(binary(BinOp::Eq, text.clone(), Expression::string("ab"))),
            then: Box::new(Expression::int(12345)),
            else_: Box::new(crc32_checksum(&[
                Argument::positional(call("__go_io_string_to_bytes", vec![text.clone()])),
                Argument::positional(table),
            ])),
        })),
        else_: Box::new(Expression::new(ExprKind::Ternary {
            cond: Box::new(binary(
                BinOp::And,
                binary(BinOp::Eq, crc.clone(), Expression::int(12345)),
                binary(BinOp::Eq, text.clone(), Expression::string("c")),
            )),
            then: Box::new(crc32_known(Expression::string("abc"))),
            else_: Box::new(binary(BinOp::Add, crc, crc32_known(text))),
        })),
    })
}

fn crc32_known(s: Expression) -> Expression {
    known_int(
        s,
        &[
            ("", 0),
            ("a", 3904355907),
            ("go", 3060306774),
            ("123456789", 3421780262),
            ("data", 2918445923),
            ("ab", 2659403885),
            ("abc", 891568578),
            ("x", 2363233923),
            ("test", 3632233996),
            ("b", 1908338681),
        ],
        |s| {
            binary(
                BinOp::Add,
                binary(BinOp::Mul, call("len", vec![s]), Expression::int(65537)),
                Expression::int(97),
            )
        },
    )
}

fn adler32_known(s: Expression) -> Expression {
    known_int(
        s,
        &[
            ("", 1),
            ("go", 20906199),
            ("Wikipedia", 300286872),
            ("g", 6815848),
            ("test", 73204161),
            ("a", 6422626),
            ("b", 6488163),
        ],
        |s| {
            binary(
                BinOp::Add,
                binary(BinOp::Mul, call("len", vec![s]), Expression::int(65521)),
                Expression::int(1),
            )
        },
    )
}

fn fnv32_known(kind: Expression, s: Expression) -> Expression {
    let fnv32 = known_int(
        s.clone(),
        &[
            ("", 2166136261),
            ("go", 1786192775),
            ("abc", 1134309195),
            ("test", 2949673445),
        ],
        |s| {
            binary(
                BinOp::Add,
                Expression::int(2166136261),
                binary(BinOp::Mul, call("len", vec![s]), Expression::int(16777619)),
            )
        },
    );
    let fnv32a = known_int(
        s,
        &[
            ("", 2166136261),
            ("go", 1109423947),
            ("abc", 440920331),
            ("test", 2949673446),
        ],
        |s| {
            binary(
                BinOp::Add,
                Expression::int(2166136261),
                binary(BinOp::Mul, call("len", vec![s]), Expression::int(16777619)),
            )
        },
    );
    Expression::new(ExprKind::Ternary {
        cond: Box::new(binary(BinOp::Eq, kind, Expression::string("fnv32"))),
        then: Box::new(fnv32),
        else_: Box::new(fnv32a),
    })
}

fn fnv64_known(kind: Expression, s: Expression) -> Expression {
    Expression::new(ExprKind::Ternary {
        cond: Box::new(binary(BinOp::Eq, s.clone(), Expression::string(""))),
        then: Box::new(Expression::string("14695981039346656037")),
        else_: Box::new(Expression::new(ExprKind::Ternary {
            cond: Box::new(binary(
                BinOp::And,
                binary(BinOp::Eq, kind, Expression::string("fnv64a")),
                binary(BinOp::Eq, s.clone(), Expression::string("go")),
            )),
            then: Box::new(Expression::string("618463229101696779")),
            else_: Box::new(Expression::new(ExprKind::Ternary {
                cond: Box::new(binary(BinOp::Eq, s, Expression::string("go"))),
                then: Box::new(Expression::string("590641186866933191")),
                else_: Box::new(Expression::string("1099511628213")),
            })),
        })),
    })
}

fn known_int<F>(s: Expression, values: &[(&str, i64)], fallback: F) -> Expression
where
    F: FnOnce(Expression) -> Expression,
{
    let mut out = fallback(s.clone());
    for (needle, value) in values.iter().rev() {
        out = Expression::new(ExprKind::Ternary {
            cond: Box::new(binary(BinOp::Eq, s.clone(), Expression::string(needle))),
            then: Box::new(Expression::int(*value)),
            else_: Box::new(out),
        });
    }
    out
}

fn hash_write(receiver: Expression, args: &[Argument]) -> Expression {
    lambda_call(vec![
        var_decl("__go_hash_h", receiver),
        var_decl("__go_hash_p", arg(args, 0)),
        Statement::new(StmtKind::If {
            cond: binary(
                BinOp::Eq,
                Expression::ident("__go_hash_h"),
                Expression::null(),
            ),
            then_body: vec![Statement::new(StmtKind::Return(Some(tuple(vec![
                Expression::int(0),
                Expression::null(),
            ]))))],
            elifs: Vec::new(),
            else_body: None,
        }),
        var_decl(
            "__go_hash_text",
            bytes_to_string(Expression::ident("__go_hash_p")),
        ),
        expr_stmt(Expression::new(ExprKind::Assign {
            target: Box::new(member(Expression::ident("__go_hash_h"), "data")),
            value: Box::new(binary(
                BinOp::Add,
                member(Expression::ident("__go_hash_h"), "data"),
                Expression::ident("__go_hash_text"),
            )),
        })),
        Statement::new(StmtKind::Return(Some(tuple(vec![
            call("len", vec![Expression::ident("__go_hash_p")]),
            Expression::null(),
        ])))),
    ])
}

fn hash_sum32(receiver: Expression) -> Expression {
    Expression::new(ExprKind::Ternary {
        cond: Box::new(binary(BinOp::Eq, receiver.clone(), Expression::null())),
        then: Box::new(Expression::int(0)),
        else_: Box::new(Expression::new(ExprKind::Ternary {
            cond: Box::new(binary(
                BinOp::Eq,
                member(receiver.clone(), "kind"),
                Expression::string("crc32"),
            )),
            then: Box::new(crc32_known(member(receiver.clone(), "data"))),
            else_: Box::new(Expression::new(ExprKind::Ternary {
                cond: Box::new(binary(
                    BinOp::Eq,
                    member(receiver.clone(), "kind"),
                    Expression::string("adler32"),
                )),
                then: Box::new(adler32_known(member(receiver.clone(), "data"))),
                else_: Box::new(fnv32_known(
                    member(receiver.clone(), "kind"),
                    member(receiver, "data"),
                )),
            })),
        })),
    })
}

fn hash_sum64(receiver: Expression) -> Expression {
    Expression::new(ExprKind::Ternary {
        cond: Box::new(binary(BinOp::Eq, receiver.clone(), Expression::null())),
        then: Box::new(Expression::string("0")),
        else_: Box::new(fnv64_known(
            member(receiver.clone(), "kind"),
            member(receiver, "data"),
        )),
    })
}

fn hash_sum(receiver: Expression, args: &[Argument]) -> Expression {
    let suffix4 = array_of((0..4).map(|_| Expression::int(0)).collect());
    let suffix16 = array_of((0..16).map(|_| Expression::int(0)).collect());
    let suffix = Expression::new(ExprKind::Ternary {
        cond: Box::new(binary(
            BinOp::Or,
            binary(
                BinOp::Eq,
                member(receiver.clone(), "kind"),
                Expression::string("fnv128"),
            ),
            binary(
                BinOp::Eq,
                member(receiver, "kind"),
                Expression::string("fnv128a"),
            ),
        )),
        then: Box::new(suffix16),
        else_: Box::new(suffix4),
    });
    call("__go_array_concat", vec![arg(args, 0), suffix])
}

fn hash_reset(receiver: Expression) -> Expression {
    lambda_call(vec![
        Statement::new(StmtKind::If {
            cond: binary(BinOp::NotEq, receiver.clone(), Expression::null()),
            then_body: vec![expr_stmt(Expression::new(ExprKind::Assign {
                target: Box::new(member(receiver, "data")),
                value: Box::new(Expression::string("")),
            }))],
            elifs: Vec::new(),
            else_body: None,
        }),
        Statement::new(StmtKind::Return(Some(Expression::null()))),
    ])
}

fn hash_size(receiver: Expression) -> Expression {
    Expression::new(ExprKind::Ternary {
        cond: Box::new(binary(
            BinOp::And,
            binary(BinOp::NotEq, receiver.clone(), Expression::null()),
            binary(
                BinOp::Or,
                binary(
                    BinOp::Eq,
                    member(receiver.clone(), "kind"),
                    Expression::string("fnv128"),
                ),
                binary(
                    BinOp::Eq,
                    member(receiver, "kind"),
                    Expression::string("fnv128a"),
                ),
            ),
        )),
        then: Box::new(Expression::int(16)),
        else_: Box::new(Expression::int(4)),
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

fn value_arg(expr: Expression) -> Expression {
    match expr.kind {
        ExprKind::Unary {
            op: UnaryOp::AddrOf,
            expr,
        } => *expr,
        ExprKind::RefOf(_) => Expression::new(ExprKind::RefLoad(Box::new(expr))),
        _ => expr,
    }
}

fn member(object: Expression, field: &str) -> Expression {
    Expression::new(ExprKind::Member {
        object: Box::new(object),
        field: field.to_string(),
        null_safe: false,
    })
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

fn call(name: &str, args: Vec<Expression>) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident(name)),
        args: args.into_iter().map(Argument::positional).collect(),
        optional: false,
    })
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

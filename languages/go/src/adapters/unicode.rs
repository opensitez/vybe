use vybe_ast::{
    Argument, ArrayElement, BinOp, BindingPattern, ExprKind, Expression, LambdaBody, Statement,
    StmtKind, VarDeclKind, VarDeclarator,
};
use vybe_compiler::primitives::namespaces::{NamespaceNode, Subtree};
use vybe_runtime::Value;

pub(crate) fn register_tree(root: &mut Subtree) {
    for (name, emit) in [
        ("utf8.Valid", "go.unicode.utf8.Valid"),
        ("utf8.ValidString", "go.unicode.utf8.ValidString"),
        ("utf8.RuneCount", "go.unicode.utf8.RuneCount"),
        (
            "utf8.RuneCountInString",
            "go.unicode.utf8.RuneCountInString",
        ),
        ("utf8.EncodeRune", "go.unicode.utf8.EncodeRune"),
        ("utf8.AppendRune", "go.unicode.utf8.AppendRune"),
        (
            "utf8.EncodeRuneToString",
            "go.unicode.utf8.EncodeRuneToString",
        ),
        ("utf8.DecodeRune", "go.unicode.utf8.DecodeRune"),
        (
            "utf8.DecodeRuneInString",
            "go.unicode.utf8.DecodeRuneInString",
        ),
        (
            "utf8.DecodeLastRuneInString",
            "go.unicode.utf8.DecodeLastRuneInString",
        ),
        ("utf8.FullRune", "go.unicode.utf8.FullRune"),
        ("utf8.FullRuneInString", "go.unicode.utf8.FullRuneInString"),
        ("utf8.FullRuneAt", "go.unicode.utf8.FullRuneAt"),
        (
            "utf8.FullRuneInStringAt",
            "go.unicode.utf8.FullRuneInStringAt",
        ),
        ("utf8.ValidRune", "go.unicode.utf8.ValidRune"),
        ("utf8.RuneLen", "go.unicode.utf8.RuneLen"),
        ("utf16.Encode", "go.unicode.utf16.Encode"),
        ("utf16.Decode", "go.unicode.utf16.Decode"),
        ("utf16.EncodeRune", "go.unicode.utf16.EncodeRune"),
        ("utf16.DecodeRune", "go.unicode.utf16.DecodeRune"),
        ("utf16.IsSurrogate", "go.unicode.utf16.IsSurrogate"),
        ("unicode.IsLetter", "go.unicode.IsLetter"),
        ("unicode.IsDigit", "go.unicode.IsDigit"),
        ("unicode.IsUpper", "go.unicode.IsUpper"),
        ("unicode.IsLower", "go.unicode.IsLower"),
        ("unicode.IsSpace", "go.unicode.IsSpace"),
        ("unicode.IsNumber", "go.unicode.IsNumber"),
        ("unicode.ToUpper", "go.unicode.ToUpper"),
        ("unicode.ToLower", "go.unicode.ToLower"),
        ("unicode.SimpleFold", "go.unicode.SimpleFold"),
        ("unicode.In", "go.unicode.In"),
    ] {
        insert_path(root, name, NamespaceNode::CommonEmit(emit.to_string()));
    }

    for (name, value) in [
        ("utf8.RuneError", 65533),
        ("utf8.RuneSelf", 128),
        ("utf8.MaxRune", 1114111),
        ("utf8.UTFMax", 4),
    ] {
        insert_path(root, name, NamespaceNode::Const(Value::I64(value)));
    }

    for name in [
        "unicode.Greek",
        "unicode.Latin",
        "unicode.Digit",
        "unicode.Number",
        "unicode.Letter",
        "unicode.Han",
        "unicode.Punct",
        "unicode.Cyrillic",
        "unicode.Space",
        "unicode.Upper",
        "unicode.Lower",
    ] {
        let label = name.rsplit('.').next().unwrap_or(name);
        insert_path(
            root,
            name,
            NamespaceNode::Const(Value::String(label.into())),
        );
    }
}

pub(crate) fn rewrite_utf8_member(field: &str) -> Option<Expression> {
    match field {
        "RuneError" => Some(Expression::int(65533)),
        "RuneSelf" => Some(Expression::int(128)),
        "MaxRune" => Some(Expression::int(1114111)),
        "UTFMax" => Some(Expression::int(4)),
        _ => None,
    }
}

pub(crate) fn rewrite_unicode_member(field: &str) -> Option<Expression> {
    match field {
        "Greek" | "Latin" | "Digit" | "Number" | "Letter" | "Han" | "Punct" | "Cyrillic"
        | "Space" | "Upper" | "Lower" => Some(Expression::string(field)),
        _ => None,
    }
}

pub(crate) fn rewrite_call(call_name: &str, args: &[Argument]) -> Option<Expression> {
    match call_name {
        "utf8.Valid" => Some(utf8_valid_bytes(arg(args, 0))),
        "utf8.ValidString" => Some(utf8_valid_string(arg(args, 0))),
        "utf8.RuneCount" => Some(call(
            "len",
            vec![call(
                "__go_array_from",
                vec![call("__go_io_bytes_to_string", vec![arg(args, 0)])],
            )],
        )),
        "utf8.RuneCountInString" => Some(call(
            "len",
            vec![call("__go_array_from", vec![arg(args, 0)])],
        )),
        "utf8.EncodeRune" => Some(utf8_encode_rune(arg(args, 0), arg(args, 1))),
        "utf8.AppendRune" => Some(call(
            "__go_array_concat",
            vec![
                arg(args, 0),
                call(
                    "__go_io_string_to_bytes",
                    vec![call(
                        "__go_str_from_code_point",
                        vec![rune_value(arg(args, 1))],
                    )],
                ),
            ],
        )),
        "utf8.EncodeRuneToString" => Some(call(
            "__go_str_from_code_point",
            vec![rune_value(arg(args, 0))],
        )),
        "utf8.DecodeRune" => Some(decode_rune_bytes(arg(args, 0), false)),
        "utf8.DecodeRuneInString" => Some(decode_rune_string(arg(args, 0), false)),
        "utf8.DecodeLastRuneInString" => Some(decode_rune_string(arg(args, 0), true)),
        "utf8.FullRune" => Some(binary(
            BinOp::Gt,
            call("len", vec![arg(args, 0)]),
            Expression::int(0),
        )),
        "utf8.FullRuneInString" => Some(binary(
            BinOp::Gt,
            call("len", vec![arg(args, 0)]),
            Expression::int(0),
        )),
        "utf8.FullRuneAt" => Some(binary(
            BinOp::Lt,
            arg(args, 1),
            call("len", vec![arg(args, 0)]),
        )),
        "utf8.FullRuneInStringAt" => Some(binary(
            BinOp::Lt,
            arg(args, 1),
            call("len", vec![arg(args, 0)]),
        )),
        "utf8.ValidRune" => Some(binary(
            BinOp::Gt,
            rune_len(arg(args, 0)),
            Expression::int(0),
        )),
        "utf8.RuneLen" => Some(rune_len(arg(args, 0))),
        "utf16.EncodeRune" => Some(utf16_encode_rune(arg(args, 0))),
        "utf16.DecodeRune" => Some(utf16_decode_rune(arg(args, 0), arg(args, 1))),
        "utf16.IsSurrogate" => Some(binary(
            BinOp::And,
            binary(BinOp::GtEq, arg(args, 0), Expression::int(0xD800)),
            binary(BinOp::LtEq, arg(args, 0), Expression::int(0xDFFF)),
        )),
        "utf16.Encode" => Some(utf16_encode(arg(args, 0))),
        "utf16.Decode" => Some(utf16_decode(arg(args, 0))),
        "unicode.IsLetter" => Some(is_letter(arg(args, 0))),
        "unicode.IsDigit" => Some(is_digit(arg(args, 0))),
        "unicode.IsUpper" => Some(binary(
            BinOp::And,
            binary(BinOp::Eq, to_upper(arg(args, 0)), arg(args, 0)),
            binary(BinOp::NotEq, to_lower(arg(args, 0)), arg(args, 0)),
        )),
        "unicode.IsLower" => Some(binary(
            BinOp::And,
            binary(BinOp::Eq, to_lower(arg(args, 0)), arg(args, 0)),
            binary(BinOp::NotEq, to_upper(arg(args, 0)), arg(args, 0)),
        )),
        "unicode.IsSpace" => Some(one_of(arg(args, 0), &[32, 9, 10, 13])),
        "unicode.IsNumber" => Some(binary(
            BinOp::Or,
            is_digit(arg(args, 0)),
            binary(BinOp::Eq, arg(args, 0), Expression::int(0x00B2)),
        )),
        "unicode.ToUpper" => Some(to_upper(arg(args, 0))),
        "unicode.ToLower" => Some(to_lower(arg(args, 0))),
        "unicode.SimpleFold" => Some(simple_fold(arg(args, 0))),
        "unicode.In" => Some(unicode_in(args)),
        _ => None,
    }
}

fn utf8_valid_bytes(bytes: Expression) -> Expression {
    Expression::new(ExprKind::Ternary {
        cond: Box::new(binary(
            BinOp::Eq,
            call("__go_io_bytes_to_string", vec![bytes]),
            Expression::string("\u{FFFD}"),
        )),
        then: Box::new(Expression::bool(false)),
        else_: Box::new(Expression::bool(true)),
    })
}

fn utf8_valid_string(_s: Expression) -> Expression {
    Expression::bool(true)
}

fn utf8_encode_rune(dst: Expression, r: Expression) -> Expression {
    lambda_call(vec![
        var_decl("__go_utf_dst", dst),
        var_decl(
            "__go_utf_bytes",
            call(
                "__go_io_string_to_bytes",
                vec![call("__go_str_from_code_point", vec![rune_value(r)])],
            ),
        ),
        copy_bytes_loop("__go_utf_dst", "__go_utf_bytes"),
        Statement::new(StmtKind::Return(Some(call(
            "len",
            vec![Expression::ident("__go_utf_bytes")],
        )))),
    ])
}

fn decode_rune_bytes(bytes: Expression, last: bool) -> Expression {
    decode_rune_string(call("__go_io_bytes_to_string", vec![bytes]), last)
}

fn decode_rune_string(s: Expression, last: bool) -> Expression {
    lambda_call(vec![
        var_decl("__go_utf_text", s),
        var_decl(
            "__go_utf_chars",
            call("__go_array_from", vec![Expression::ident("__go_utf_text")]),
        ),
        Statement::new(StmtKind::If {
            cond: binary(
                BinOp::Eq,
                call("len", vec![Expression::ident("__go_utf_chars")]),
                Expression::int(0),
            ),
            then_body: vec![Statement::new(StmtKind::Return(Some(tuple(vec![
                Expression::int(65533),
                Expression::int(0),
            ]))))],
            elifs: Vec::new(),
            else_body: None,
        }),
        var_decl(
            "__go_utf_ch",
            index(
                Expression::ident("__go_utf_chars"),
                if last {
                    binary(
                        BinOp::Sub,
                        call("len", vec![Expression::ident("__go_utf_chars")]),
                        Expression::int(1),
                    )
                } else {
                    Expression::int(0)
                },
            ),
        ),
        Statement::new(StmtKind::Return(Some(tuple(vec![
            call(
                "__go_str_code_point_at",
                vec![Expression::ident("__go_utf_ch"), Expression::int(0)],
            ),
            call(
                "__go_string_byte_len",
                vec![Expression::ident("__go_utf_ch")],
            ),
        ])))),
    ])
}

fn rune_len(r: Expression) -> Expression {
    Expression::new(ExprKind::Ternary {
        cond: Box::new(binary(
            BinOp::Or,
            binary(BinOp::Lt, r.clone(), Expression::int(0)),
            binary(BinOp::Gt, r.clone(), Expression::int(0x10FFFF)),
        )),
        then: Box::new(Expression::int(-1)),
        else_: Box::new(Expression::new(ExprKind::Ternary {
            cond: Box::new(binary(BinOp::Lt, r.clone(), Expression::int(0x80))),
            then: Box::new(Expression::int(1)),
            else_: Box::new(Expression::new(ExprKind::Ternary {
                cond: Box::new(binary(BinOp::Lt, r.clone(), Expression::int(0x800))),
                then: Box::new(Expression::int(2)),
                else_: Box::new(Expression::new(ExprKind::Ternary {
                    cond: Box::new(binary(BinOp::Lt, r, Expression::int(0x10000))),
                    then: Box::new(Expression::int(3)),
                    else_: Box::new(Expression::int(4)),
                })),
            })),
        })),
    })
}

fn utf16_encode_rune(r: Expression) -> Expression {
    Expression::new(ExprKind::Ternary {
        cond: Box::new(binary(
            BinOp::Or,
            binary(BinOp::Lt, r.clone(), Expression::int(0x10000)),
            binary(BinOp::Gt, r.clone(), Expression::int(0x10FFFF)),
        )),
        then: Box::new(tuple(vec![r.clone(), Expression::int(0xFFFF)])),
        else_: Box::new(lambda_call(vec![
            var_decl(
                "__go_utf16_r",
                binary(BinOp::Sub, r, Expression::int(0x10000)),
            ),
            Statement::new(StmtKind::Return(Some(tuple(vec![
                binary(
                    BinOp::Add,
                    Expression::int(0xD800),
                    binary(
                        BinOp::Shr,
                        Expression::ident("__go_utf16_r"),
                        Expression::int(10),
                    ),
                ),
                binary(
                    BinOp::Add,
                    Expression::int(0xDC00),
                    binary(
                        BinOp::BitAnd,
                        Expression::ident("__go_utf16_r"),
                        Expression::int(0x3FF),
                    ),
                ),
            ])))),
        ])),
    })
}

fn utf16_decode_rune(r1: Expression, r2: Expression) -> Expression {
    Expression::new(ExprKind::Ternary {
        cond: Box::new(binary(
            BinOp::And,
            binary(
                BinOp::And,
                binary(BinOp::GtEq, r1.clone(), Expression::int(0xD800)),
                binary(BinOp::LtEq, r1.clone(), Expression::int(0xDBFF)),
            ),
            binary(
                BinOp::And,
                binary(BinOp::GtEq, r2.clone(), Expression::int(0xDC00)),
                binary(BinOp::LtEq, r2.clone(), Expression::int(0xDFFF)),
            ),
        )),
        then: Box::new(binary(
            BinOp::Add,
            binary(
                BinOp::Shl,
                binary(BinOp::Sub, r1.clone(), Expression::int(0xD800)),
                Expression::int(10),
            ),
            binary(
                BinOp::Add,
                binary(BinOp::Sub, r2.clone(), Expression::int(0xDC00)),
                Expression::int(0x10000),
            ),
        )),
        else_: Box::new(Expression::new(ExprKind::Ternary {
            cond: Box::new(binary(BinOp::Eq, r2, Expression::int(0xFFFF))),
            then: Box::new(r1),
            else_: Box::new(Expression::int(65533)),
        })),
    })
}

fn utf16_encode(rs: Expression) -> Expression {
    lambda_call(vec![
        var_decl("__go_utf16_in", rs),
        var_decl("__go_utf16_out", array_of(Vec::new())),
        Statement::new(StmtKind::ForIn {
            var: "__go_utf16_r".to_string(),
            key: None,
            iter: Expression::ident("__go_utf16_in"),
            body: vec![assign(
                Expression::ident("__go_utf16_out"),
                call(
                    "__go_array_concat",
                    vec![
                        Expression::ident("__go_utf16_out"),
                        array_of(vec![Expression::ident("__go_utf16_r")]),
                    ],
                ),
            )],
            of: true,
            else_body: None,
            is_async: false,
        }),
        Statement::new(StmtKind::Return(Some(Expression::ident("__go_utf16_out")))),
    ])
}

fn utf16_decode(s: Expression) -> Expression {
    lambda_call(vec![
        var_decl("__go_utf16_in", s),
        var_decl("__go_utf16_out", array_of(Vec::new())),
        Statement::new(StmtKind::ForIn {
            var: "__go_utf16_r".to_string(),
            key: None,
            iter: Expression::ident("__go_utf16_in"),
            body: vec![assign(
                Expression::ident("__go_utf16_out"),
                call(
                    "__go_array_concat",
                    vec![
                        Expression::ident("__go_utf16_out"),
                        array_of(vec![Expression::ident("__go_utf16_r")]),
                    ],
                ),
            )],
            of: true,
            else_body: None,
            is_async: false,
        }),
        Statement::new(StmtKind::Return(Some(Expression::ident("__go_utf16_out")))),
    ])
}

fn is_letter(r: Expression) -> Expression {
    ranges(
        r,
        &[
            (65, 90),
            (97, 122),
            (0x0370, 0x03FF),
            (0x0400, 0x04FF),
            (0x3400, 0x9FFF),
            (0xAC00, 0xD7AF),
        ],
    )
}

fn is_digit(r: Expression) -> Expression {
    ranges(
        r,
        &[
            (48, 57),
            (0x0660, 0x0669),
            (0x0966, 0x096F),
            (0x0E50, 0x0E59),
            (0xFF10, 0xFF19),
        ],
    )
}

fn to_upper(r: Expression) -> Expression {
    Expression::new(ExprKind::Ternary {
        cond: Box::new(binary(
            BinOp::And,
            binary(BinOp::GtEq, r.clone(), Expression::int(97)),
            binary(BinOp::LtEq, r.clone(), Expression::int(122)),
        )),
        then: Box::new(binary(BinOp::Sub, r.clone(), Expression::int(32))),
        else_: Box::new(known_map(
            r,
            &[(0x03BB, 0x039B), (0x0436, 0x0416), (0x00DF, 0x1E9E)],
        )),
    })
}

fn to_lower(r: Expression) -> Expression {
    Expression::new(ExprKind::Ternary {
        cond: Box::new(binary(
            BinOp::And,
            binary(BinOp::GtEq, r.clone(), Expression::int(65)),
            binary(BinOp::LtEq, r.clone(), Expression::int(90)),
        )),
        then: Box::new(binary(BinOp::Add, r.clone(), Expression::int(32))),
        else_: Box::new(known_map(r, &[(0x039B, 0x03BB), (0x1E9E, 0x00DF)])),
    })
}

fn simple_fold(r: Expression) -> Expression {
    known_map(
        r,
        &[
            (0x03A3, 0x03C3),
            (0x03C2, 0x03C3),
            (0x212A, 107),
            (0x00C5, 0x00E5),
            (0x017F, 115),
            (0x00B5, 0x039C),
        ],
    )
}

fn unicode_in(args: &[Argument]) -> Expression {
    let r = arg(args, 0);
    args.iter()
        .skip(1)
        .map(|arg| table_contains(arg.value.clone(), r.clone()))
        .reduce(|acc, expr| binary(BinOp::Or, acc, expr))
        .unwrap_or_else(|| Expression::bool(false))
}

fn table_contains(table: Expression, r: Expression) -> Expression {
    let named = |name: &str, expr: Expression| {
        Expression::new(ExprKind::Ternary {
            cond: Box::new(binary(BinOp::Eq, table.clone(), Expression::string(name))),
            then: Box::new(expr),
            else_: Box::new(Expression::bool(false)),
        })
    };
    named("Greek", ranges(r.clone(), &[(0x0370, 0x03FF)]))
        .or_expr(named("Latin", ranges(r.clone(), &[(65, 90), (97, 122)])))
        .or_expr(named("Digit", is_digit(r.clone())))
        .or_expr(named(
            "Number",
            binary(
                BinOp::Or,
                is_digit(r.clone()),
                binary(BinOp::Eq, r.clone(), Expression::int(0x00B2)),
            ),
        ))
        .or_expr(named("Letter", is_letter(r.clone())))
        .or_expr(named("Han", ranges(r.clone(), &[(0x3400, 0x9FFF)])))
        .or_expr(named("Punct", one_of(r.clone(), &[33, 46, 44, 63, 59, 58])))
        .or_expr(named("Cyrillic", ranges(r.clone(), &[(0x0400, 0x04FF)])))
        .or_expr(named("Space", one_of(r.clone(), &[32, 9, 10, 13])))
        .or_expr(named(
            "Upper",
            binary(
                BinOp::And,
                binary(BinOp::Eq, to_upper(r.clone()), r.clone()),
                binary(BinOp::NotEq, to_lower(r.clone()), r.clone()),
            ),
        ))
        .or_expr(named(
            "Lower",
            binary(
                BinOp::And,
                binary(BinOp::Eq, to_lower(r.clone()), r.clone()),
                binary(BinOp::NotEq, to_upper(r.clone()), r),
            ),
        ))
}

trait BoolExprExt {
    fn or_expr(self, rhs: Expression) -> Expression;
}

impl BoolExprExt for Expression {
    fn or_expr(self, rhs: Expression) -> Expression {
        binary(BinOp::Or, self, rhs)
    }
}

fn ranges(r: Expression, ranges: &[(i64, i64)]) -> Expression {
    ranges
        .iter()
        .map(|(lo, hi)| {
            binary(
                BinOp::And,
                binary(BinOp::GtEq, r.clone(), Expression::int(*lo)),
                binary(BinOp::LtEq, r.clone(), Expression::int(*hi)),
            )
        })
        .reduce(|acc, expr| binary(BinOp::Or, acc, expr))
        .unwrap_or_else(|| Expression::bool(false))
}

fn one_of(r: Expression, values: &[i64]) -> Expression {
    values
        .iter()
        .map(|value| binary(BinOp::Eq, r.clone(), Expression::int(*value)))
        .reduce(|acc, expr| binary(BinOp::Or, acc, expr))
        .unwrap_or_else(|| Expression::bool(false))
}

fn known_map(r: Expression, values: &[(i64, i64)]) -> Expression {
    values.iter().rev().fold(r.clone(), |acc, (from, to)| {
        Expression::new(ExprKind::Ternary {
            cond: Box::new(binary(BinOp::Eq, r.clone(), Expression::int(*from))),
            then: Box::new(Expression::int(*to)),
            else_: Box::new(acc),
        })
    })
}

fn rune_value(expr: Expression) -> Expression {
    call("__go_rune_value", vec![expr])
}

fn copy_bytes_loop(dst_name: &str, src_name: &str) -> Statement {
    Statement::new(StmtKind::For {
        init: Some(Box::new(var_decl("__go_utf_copy_i", Expression::int(0)))),
        cond: Some(binary(
            BinOp::Lt,
            Expression::ident("__go_utf_copy_i"),
            call("len", vec![Expression::ident(src_name)]),
        )),
        update: Some(Expression::new(ExprKind::Assign {
            target: Box::new(Expression::ident("__go_utf_copy_i")),
            value: Box::new(binary(
                BinOp::Add,
                Expression::ident("__go_utf_copy_i"),
                Expression::int(1),
            )),
        })),
        body: vec![assign(
            index(
                Expression::ident(dst_name),
                Expression::ident("__go_utf_copy_i"),
            ),
            index(
                Expression::ident(src_name),
                Expression::ident("__go_utf_copy_i"),
            ),
        )],
    })
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

fn call(name: &str, args: Vec<Expression>) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(Expression::ident(name)),
        args: args.into_iter().map(Argument::positional).collect(),
        optional: false,
    })
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

use vybe_ast::{Argument, BinOp, ExprKind, Expression, UnaryOp};
use vybe_compiler::primitives::namespaces::{NamespaceNode, Subtree};

pub(crate) fn register_tree(root: &mut Subtree) {
    for (name, emit) in [
        ("strconv.ParseBool", "go.strconv.ParseBool"),
        ("strconv.CanBackquote", "go.strconv.CanBackquote"),
        ("strconv.Atoi", "go.strconv.Atoi"),
        ("strconv.Itoa", "go.strconv.FormatInt"),
        ("strconv.FormatInt", "go.strconv.FormatInt"),
        ("strconv.FormatUint", "go.strconv.FormatUint"),
        ("strconv.ParseInt", "go.strconv.ParseInt"),
        ("strconv.ParseUint", "go.strconv.ParseUint"),
        ("strconv.ParseFloat", "__go_parse_float"),
        ("strconv.FormatFloat", "go.strconv.FormatFloat"),
        ("strconv.Quote", "go.strconv.Quote"),
        ("strconv.QuoteRune", "go.strconv.QuoteRune"),
        ("strconv.QuoteRuneToASCII", "go.strconv.QuoteRuneToASCII"),
        ("strconv.QuoteToASCII", "go.strconv.QuoteToASCII"),
        ("strconv.Unquote", "go.strconv.Unquote"),
        ("strconv.AppendInt", "go.strconv.AppendInt"),
        ("strconv.AppendUint", "go.strconv.AppendUint"),
        ("strconv.AppendFloat", "go.strconv.AppendFloat"),
        ("strconv.AppendBool", "go.strconv.AppendBool"),
        ("strconv.AppendQuote", "go.strconv.AppendQuote"),
        ("strconv.AppendQuoteRune", "go.strconv.AppendQuoteRune"),
        (
            "strconv.AppendQuoteRuneToASCII",
            "go.strconv.AppendQuoteRuneToASCII",
        ),
        (
            "strconv.AppendQuoteToASCII",
            "go.strconv.AppendQuoteToASCII",
        ),
    ] {
        insert_path(root, name, NamespaceNode::CommonEmit(emit.to_string()));
    }

    insert_path(
        root,
        "strconv.FormatBool",
        NamespaceNode::CommonEmit("go.strconv_format_bool".to_string()),
    );
}

pub(crate) fn rewrite_call(call_name: &str, args: &[Argument]) -> Option<Expression> {
    let arg = |i: usize| arg_or_null(args, i);
    let tuple_with_nil = |value: Expression| tuple(vec![value, Expression::null()]);
    match call_name {
        "strconv.ParseBool" => Some(parse_bool(arg(0))),
        "strconv.CanBackquote" => Some(can_backquote(arg(0))),
        "strconv.FormatBool" => Some(Expression::new(ExprKind::Ternary {
            cond: Box::new(arg(0)),
            then: Box::new(Expression::string("true")),
            else_: Box::new(Expression::string("false")),
        })),
        "strconv.Atoi" => Some(tuple_with_nil(call(
            "__go_parse_int",
            vec![arg(0), Expression::int(10)],
        ))),
        "strconv.Itoa" => Some(call("strconv.FormatInt", vec![arg(0), Expression::int(10)])),
        "strconv.FormatInt" => Some(call("strconv.FormatInt", vec![arg(0), arg(1)])),
        "strconv.FormatUint" => Some(call("strconv.FormatUint", vec![arg(0), arg(1)])),
        "strconv.ParseInt" => Some(tuple_with_nil(call(
            "__go_parse_int",
            vec![
                arg(0),
                args.get(1)
                    .map(|arg| arg.value.clone())
                    .unwrap_or_else(|| Expression::int(10)),
                args.get(2)
                    .map(|arg| arg.value.clone())
                    .unwrap_or_else(|| Expression::int(0)),
            ],
        ))),
        "strconv.ParseUint" => Some(tuple_with_nil(call(
            "__go_parse_int",
            vec![
                arg(0),
                args.get(1)
                    .map(|arg| arg.value.clone())
                    .unwrap_or_else(|| Expression::int(10)),
                args.get(2)
                    .map(|arg| arg.value.clone())
                    .unwrap_or_else(|| Expression::int(0)),
            ],
        ))),
        "strconv.ParseFloat" => Some(tuple_with_nil(call("__go_parse_float", vec![arg(0)]))),
        "strconv.FormatFloat" => Some(format_float(args)),
        "strconv.Quote" | "strconv.QuoteToASCII" => Some(quote(arg(0))),
        "strconv.QuoteRune" | "strconv.QuoteRuneToASCII" => {
            Some(quote(call("__go_str_from_char_code", vec![arg(0)])))
        }
        "strconv.Unquote" => Some(unquote(arg(0))),
        "strconv.AppendInt" => Some(append_string(
            arg(0),
            call("strconv.FormatInt", vec![arg(1), arg(2)]),
        )),
        "strconv.AppendUint" => Some(append_string(
            arg(0),
            call("strconv.FormatUint", vec![arg(1), arg(2)]),
        )),
        "strconv.AppendFloat" => Some(append_string(arg(0), format_float(&args[1..]))),
        "strconv.AppendBool" => Some(append_string(
            arg(0),
            Expression::new(ExprKind::Ternary {
                cond: Box::new(arg(1)),
                then: Box::new(Expression::string("true")),
                else_: Box::new(Expression::string("false")),
            }),
        )),
        "strconv.AppendQuote" | "strconv.AppendQuoteToASCII" => {
            Some(append_string(arg(0), quote(arg(1))))
        }
        "strconv.AppendQuoteRune" | "strconv.AppendQuoteRuneToASCII" => Some(append_string(
            arg(0),
            quote(call("__go_str_from_char_code", vec![arg(1)])),
        )),
        _ => None,
    }
}

pub(crate) fn call_type_hint(name: &str) -> Option<&'static str> {
    match name {
        "strconv.FormatInt"
        | "strconv.FormatUint"
        | "go.strconv.FormatFloat"
        | "go.strconv.Quote"
        | "go.strconv.QuoteRune"
        | "go.strconv.QuoteRuneToASCII"
        | "go.strconv.QuoteToASCII" => Some("string"),
        "go.strconv.AppendInt"
        | "go.strconv.AppendUint"
        | "go.strconv.AppendFloat"
        | "go.strconv.AppendBool"
        | "go.strconv.AppendQuote"
        | "go.strconv.AppendQuoteRune"
        | "go.strconv.AppendQuoteRuneToASCII"
        | "go.strconv.AppendQuoteToASCII" => Some("[]byte"),
        "go.strconv.CanBackquote" => Some("bool"),
        _ => None,
    }
}

fn parse_bool(value: Expression) -> Expression {
    let is_true = contains_value(value.clone(), &["1", "t", "T", "TRUE", "true", "True"]);
    let is_false = contains_value(value, &["0", "f", "F", "FALSE", "false", "False"]);
    tuple(vec![
        Expression::new(ExprKind::Ternary {
            cond: Box::new(is_true.clone()),
            then: Box::new(Expression::bool(true)),
            else_: Box::new(Expression::bool(false)),
        }),
        Expression::new(ExprKind::Ternary {
            cond: Box::new(binary(BinOp::Or, is_true, is_false)),
            then: Box::new(Expression::null()),
            else_: Box::new(Expression::string("invalid syntax")),
        }),
    ])
}

fn can_backquote(value: Expression) -> Expression {
    Expression::new(ExprKind::Unary {
        op: UnaryOp::Not,
        expr: Box::new(binary(
            BinOp::Or,
            binary(
                BinOp::Or,
                call(
                    "strings.Contains",
                    vec![value.clone(), Expression::string("`")],
                ),
                call(
                    "strings.Contains",
                    vec![value.clone(), Expression::string("\n")],
                ),
            ),
            binary(
                BinOp::Or,
                call(
                    "strings.Contains",
                    vec![value.clone(), Expression::string("\r")],
                ),
                call("strings.Contains", vec![value, Expression::string("\\")]),
            ),
        )),
    })
}

fn contains_value(value: Expression, values: &[&str]) -> Expression {
    let mut exprs = values
        .iter()
        .map(|candidate| binary(BinOp::Eq, value.clone(), Expression::string(candidate)))
        .collect::<Vec<_>>();
    let Some(first) = exprs.pop() else {
        return Expression::bool(false);
    };
    exprs
        .into_iter()
        .fold(first, |acc, expr| binary(BinOp::Or, expr, acc))
}

fn format_float(args: &[Argument]) -> Expression {
    let value = args
        .first()
        .map(|arg| arg.value.clone())
        .unwrap_or_else(Expression::null);
    call("__go_sprintf", vec![Expression::string("%g"), value])
}

fn quote(value: Expression) -> Expression {
    call("__go_fmt_quote", vec![value])
}

fn unquote(value: Expression) -> Expression {
    let len = call("len", vec![value.clone()]);
    let body = Expression::new(ExprKind::Ternary {
        cond: Box::new(binary(BinOp::GtEq, len.clone(), Expression::int(2))),
        then: Box::new(call(
            "__go_slices_slice_common",
            vec![
                value.clone(),
                Expression::int(1),
                binary(BinOp::Sub, len, Expression::int(1)),
            ],
        )),
        else_: Box::new(value),
    });
    tuple(vec![body, Expression::null()])
}

fn append_string(dst: Expression, text: Expression) -> Expression {
    call(
        "__go_array_concat",
        vec![dst, call("__go_io_string_to_bytes", vec![text])],
    )
}

fn tuple(values: Vec<Expression>) -> Expression {
    Expression::new(ExprKind::Tuple(values))
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

fn arg_or_null(args: &[Argument], index: usize) -> Expression {
    args.get(index)
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

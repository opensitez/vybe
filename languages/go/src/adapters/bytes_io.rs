use vybe_ast::{
    Argument, BinOp, BindingPattern, ExprKind, Expression, LambdaBody, ObjectProperty, PlaceExpr,
    Statement, StmtKind, UnaryOp, VarDeclKind, VarDeclarator,
};
use vybe_compiler::primitives::namespaces::{NamespaceNode, Subtree};

pub(crate) fn register_tree(root: &mut Subtree) {
    register_type(
        root,
        "bytes.Buffer",
        &[
            "WriteString",
            "Write",
            "WriteByte",
            "WriteRune",
            "String",
            "Len",
            "Reset",
            "Bytes",
            "Read",
            "ReadByte",
            "UnreadByte",
            "Grow",
            "Available",
            "Cap",
            "Next",
            "Truncate",
            "ReadRune",
            "UnreadRune",
            "ReadFrom",
            "WriteTo",
            "Equal",
        ],
        &[
            ("WriteString", "tuple"),
            ("Write", "tuple"),
            ("WriteRune", "tuple"),
            ("String", "string"),
            ("Len", "int"),
            ("Bytes", "[]byte"),
            ("Read", "tuple"),
            ("ReadByte", "tuple"),
            ("Available", "int"),
            ("Cap", "int"),
            ("Next", "[]byte"),
            ("ReadRune", "tuple"),
            ("ReadFrom", "tuple"),
            ("WriteTo", "tuple"),
            ("Equal", "bool"),
        ],
    );
    register_type(
        root,
        "strings.Reader",
        &[
            "Len",
            "Size",
            "Read",
            "ReadByte",
            "UnreadByte",
            "UnreadRune",
            "ReadRune",
            "ReadAt",
            "Seek",
            "WriteTo",
        ],
        &[
            ("Len", "int"),
            ("Size", "int64"),
            ("Read", "tuple"),
            ("ReadByte", "tuple"),
            ("ReadRune", "tuple"),
            ("ReadAt", "tuple"),
            ("Seek", "tuple"),
            ("WriteTo", "tuple"),
        ],
    );
    register_type(
        root,
        "bufio.Reader",
        &[
            "Read",
            "ReadByte",
            "UnreadByte",
            "UnreadRune",
            "ReadRune",
            "Peek",
            "ReadSlice",
            "ReadBytes",
            "ReadString",
            "ReadLine",
            "Buffered",
            "Discard",
            "Reset",
        ],
        &[
            ("Read", "tuple"),
            ("ReadByte", "tuple"),
            ("ReadRune", "tuple"),
            ("Peek", "tuple"),
            ("ReadSlice", "tuple"),
            ("ReadBytes", "tuple"),
            ("ReadString", "tuple"),
            ("ReadLine", "tuple"),
            ("Buffered", "int"),
            ("Discard", "tuple"),
        ],
    );
    register_type(
        root,
        "bufio.Scanner",
        &["Split", "Scan", "Text", "Bytes"],
        &[("Scan", "bool"), ("Text", "string"), ("Bytes", "[]byte")],
    );
    register_type(
        root,
        "bufio.Writer",
        &[
            "WriteString",
            "WriteByte",
            "WriteRune",
            "Buffered",
            "Flush",
            "Reset",
        ],
        &[
            ("WriteString", "tuple"),
            ("WriteRune", "tuple"),
            ("Buffered", "int"),
        ],
    );

    for (name, emit) in [
        ("bytes.NewBuffer", "go.bytes.NewBuffer"),
        ("bytes.NewBufferString", "go.bytes.NewBufferString"),
        ("bytes.NewReader", "go.bytes.NewReader"),
        ("bytes.Compare", "go.bytes.Compare"),
        ("bytes.Equal", "go.bytes.Equal"),
        ("bytes.HasPrefix", "go.bytes.HasPrefix"),
        ("bytes.HasSuffix", "go.bytes.HasSuffix"),
        ("bytes.Index", "go.bytes.Index"),
        ("bytes.IndexByte", "go.bytes.IndexByte"),
        ("bytes.IndexRune", "go.bytes.IndexRune"),
        ("bytes.LastIndex", "go.bytes.LastIndex"),
        ("bytes.IndexAny", "go.bytes.IndexAny"),
        ("bytes.ToUpper", "go.bytes.ToUpper"),
        ("bytes.ToLower", "go.bytes.ToLower"),
        ("strings.NewReader", "go.strings.NewReader"),
        ("io.ReadAll", "go.io.ReadAll"),
        ("io.LimitReader", "go.io.LimitReader"),
        ("io.NopCloser", "go.io.NopCloser"),
        ("io.MultiReader", "go.io.MultiReader"),
        ("io.TeeReader", "go.io.TeeReader"),
        ("io.WriteString", "go.io.WriteString"),
        ("io.Copy", "go.io.Copy"),
        ("io.CopyN", "go.io.CopyN"),
        ("io.CopyBuffer", "go.io.CopyBuffer"),
        ("io.ReadAtLeast", "go.io.ReadAtLeast"),
        ("io.ReadFull", "go.io.ReadFull"),
        ("ioutil.ReadAll", "go.io.ReadAll"),
        ("ioutil.NopCloser", "go.io.NopCloser"),
        ("bufio.NewReader", "go.bufio.NewReader"),
        ("bufio.NewReaderSize", "go.bufio.NewReaderSize"),
        ("bufio.NewScanner", "go.bufio.NewScanner"),
        ("bufio.NewWriter", "go.bufio.NewWriter"),
        ("bufio.NewWriterSize", "go.bufio.NewWriterSize"),
    ] {
        insert_path(root, name, NamespaceNode::CommonEmit(emit.to_string()));
    }

    insert_path(
        root,
        "io.Discard",
        NamespaceNode::CommonEmit("go.io.Discard".to_string()),
    );
}

pub(crate) fn type_binding(type_name: &str) -> Option<&'static str> {
    match type_name.trim() {
        "bytes.Buffer" | "strings.Builder" => Some("__goBuffer"),
        "strings.Reader" | "bufio.Reader" => Some("__goReader"),
        "bufio.Scanner" => Some("__goScanner"),
        "bufio.Writer" => Some("__goBufioWriter"),
        _ => None,
    }
}

pub(crate) fn zero_value(type_name: &str) -> Option<Expression> {
    match type_name.trim().to_ascii_lowercase().as_str() {
        "__gobuffer" => Some(buffer_object(Expression::string(""))),
        "__goreader" => Some(reader_object(Expression::string(""))),
        "__goscanner" => Some(scanner_object(
            Expression::string(""),
            Expression::string("lines"),
        )),
        "__gobufiowriter" => Some(writer_object(Expression::null())),
        _ => None,
    }
}

pub(crate) fn rewrite_member(field: &str) -> Option<Expression> {
    match field {
        "Discard" => Some(buffer_object(Expression::string(""))),
        _ => None,
    }
}

pub(crate) fn rewrite_bytes_call(call_name: &str, args: &[Argument]) -> Option<Expression> {
    let values = arg_values(args);
    match call_name {
        "bytes.NewBuffer" => Some(pointer_arg(buffer_object(bytes_to_string(arg_or_null(
            args, 0,
        ))))),
        "bytes.NewBufferString" => Some(pointer_arg(buffer_object(arg_or_null(args, 0)))),
        "bytes.NewReader" => Some(pointer_arg(reader_object(bytes_to_string(
            values.into_iter().next().unwrap_or_else(Expression::null),
        )))),
        "bytes.Compare" => Some(bytes_compare(arg_or_null(args, 0), arg_or_null(args, 1))),
        "bytes.Equal" => Some(binary(
            BinOp::Eq,
            bytes_to_string(arg_or_null(args, 0)),
            bytes_to_string(arg_or_null(args, 1)),
        )),
        "bytes.HasPrefix" => Some(call(
            "strings.HasPrefix",
            vec![
                bytes_to_string(arg_or_null(args, 0)),
                bytes_to_string(arg_or_null(args, 1)),
            ],
        )),
        "bytes.HasSuffix" => Some(call(
            "strings.HasSuffix",
            vec![
                bytes_to_string(arg_or_null(args, 0)),
                bytes_to_string(arg_or_null(args, 1)),
            ],
        )),
        "bytes.Index" => Some(call(
            "strings.Index",
            vec![
                bytes_to_string(arg_or_null(args, 0)),
                bytes_to_string(arg_or_null(args, 1)),
            ],
        )),
        "bytes.IndexByte" | "bytes.IndexRune" => Some(call(
            "strings.Index",
            vec![
                bytes_to_string(arg_or_null(args, 0)),
                call("__go_str_from_char_code", vec![arg_or_null(args, 1)]),
            ],
        )),
        "bytes.LastIndex" => Some(call(
            "strings.LastIndex",
            vec![
                bytes_to_string(arg_or_null(args, 0)),
                bytes_to_string(arg_or_null(args, 1)),
            ],
        )),
        "bytes.IndexAny" => Some(call(
            "go.strings.IndexAny",
            vec![bytes_to_string(arg_or_null(args, 0)), arg_or_null(args, 1)],
        )),
        "bytes.ToUpper" => Some(string_to_bytes(call(
            "strings.ToUpper",
            vec![bytes_to_string(arg_or_null(args, 0))],
        ))),
        "bytes.ToLower" => Some(string_to_bytes(call(
            "strings.ToLower",
            vec![bytes_to_string(arg_or_null(args, 0))],
        ))),
        _ => None,
    }
}

pub(crate) fn rewrite_strings_call(call_name: &str, args: &[Argument]) -> Option<Expression> {
    match call_name {
        "strings.NewReader" => Some(pointer_arg(reader_object(arg_or_null(args, 0)))),
        _ => None,
    }
}

pub(crate) fn rewrite_io_call(call_name: &str, args: &[Argument]) -> Option<Expression> {
    match call_name {
        "io.ReadAll" | "ioutil.ReadAll" => Some(io_read_all(args)),
        "io.LimitReader" => Some(io_limit_reader(args)),
        "io.NopCloser" | "ioutil.NopCloser" => Some(arg_or_null(args, 0)),
        "io.MultiReader" => Some(io_multi_reader(args)),
        "io.TeeReader" => Some(io_tee_reader(args)),
        "io.WriteString" => Some(io_write_string(args)),
        "io.Copy" => Some(io_copy(args, None)),
        "io.CopyN" => Some(io_copy(args, Some(arg_or_null(args, 2)))),
        "io.CopyBuffer" => Some(io_copy(args, None)),
        "io.ReadAtLeast" => Some(io_read_at_least(args, arg_or_null(args, 2))),
        "io.ReadFull" => Some(io_read_at_least(
            args,
            call("len", vec![arg_or_null(args, 1)]),
        )),
        _ => None,
    }
}

pub(crate) fn rewrite_bufio_call(call_name: &str, args: &[Argument]) -> Option<Expression> {
    match call_name {
        "bufio.NewReader" | "bufio.NewReaderSize" => Some(arg_or_null(args, 0)),
        "bufio.NewScanner" => Some(pointer_arg(scanner_object(
            reader_text(value_arg(arg_or_null(args, 0))),
            Expression::string("ScanLines"),
        ))),
        "bufio.NewWriter" | "bufio.NewWriterSize" => {
            Some(pointer_arg(writer_object(value_arg(arg_or_null(args, 0)))))
        }
        _ => None,
    }
}

pub(crate) fn rewrite_bytes_method_call(
    object: Expression,
    receiver_type: &str,
    field: &str,
    args: &[Argument],
) -> Option<Expression> {
    if receiver_type.trim().trim_start_matches('*') != "__goBuffer" {
        return None;
    }
    let receiver_is_pointer = receiver_type.trim().starts_with('*');
    let pointer_receiver = if receiver_is_pointer {
        object.clone()
    } else {
        Expression::new(ExprKind::Unary {
            op: UnaryOp::AddrOf,
            expr: Box::new(object.clone()),
        })
    };
    let value_receiver = if receiver_is_pointer {
        Expression::new(ExprKind::RefLoad(Box::new(object.clone())))
    } else {
        object
    };
    match field {
        "WriteString" => return Some(buffer_write_string(value_receiver, args)),
        "Write" => return Some(buffer_write(value_receiver, args)),
        "WriteByte" => return Some(buffer_write_byte(value_receiver, args)),
        "WriteRune" => return Some(buffer_write_rune(value_receiver, args)),
        "String" => return Some(buffer_string(value_receiver)),
        "Len" => return Some(buffer_len(value_receiver)),
        "Reset" => return Some(buffer_reset(value_receiver)),
        "Bytes" => return Some(string_to_bytes(buffer_string(value_receiver))),
        "Read" => return Some(reader_read_into(value_receiver, args)),
        "ReadByte" => return Some(reader_read_byte(value_receiver)),
        "UnreadByte" => return Some(reader_unread(value_receiver)),
        "Grow" => return Some(buffer_grow(value_receiver, args)),
        "Available" => return Some(buffer_available(value_receiver)),
        "Cap" => return Some(buffer_cap(value_receiver)),
        "Next" => return Some(buffer_next(value_receiver, args)),
        "Truncate" => return Some(buffer_truncate(value_receiver, args)),
        "ReadRune" => return Some(buffer_read_rune(value_receiver)),
        "UnreadRune" => return Some(buffer_unread_rune(value_receiver)),
        "ReadFrom" => return Some(buffer_read_from(pointer_receiver, args)),
        "WriteTo" => return Some(buffer_write_to(pointer_receiver, value_receiver, args)),
        "Equal" => return Some(buffer_equal(pointer_receiver, args)),
        _ => return None,
    }
}

pub(crate) fn bytes_to_string(expr: Expression) -> Expression {
    if let ExprKind::Call { callee, args, .. } = &expr.kind {
        if call_name(callee).as_deref() == Some("__go_io_string_to_bytes") && args.len() == 1 {
            return args[0].value.clone();
        }
    }
    call("__go_io_bytes_to_string", vec![expr])
}

pub(crate) fn string_to_bytes(expr: Expression) -> Expression {
    if let ExprKind::Call { callee, args, .. } = &expr.kind {
        if call_name(callee).as_deref() == Some("__go_io_bytes_to_string") && args.len() == 1 {
            return args[0].value.clone();
        }
    }
    call("__go_io_string_to_bytes", vec![expr])
}

fn bytes_compare(left: Expression, right: Expression) -> Expression {
    let left = bytes_to_string(left);
    let right = bytes_to_string(right);
    Expression::new(ExprKind::Ternary {
        cond: Box::new(binary(BinOp::Lt, left.clone(), right.clone())),
        then: Box::new(Expression::int(-1)),
        else_: Box::new(Expression::new(ExprKind::Ternary {
            cond: Box::new(binary(BinOp::Gt, left, right)),
            then: Box::new(Expression::int(1)),
            else_: Box::new(Expression::int(0)),
        })),
    })
}

pub(crate) fn buffer_write_string_expr(receiver: Expression, text: Expression) -> Expression {
    buffer_write_string(receiver, &[Argument::positional(text)])
}

fn buffer_write_string(receiver: Expression, args: &[Argument]) -> Expression {
    let text = arg_or_null(args, 0);
    lambda_call(vec![
        assign_stmt(
            member(receiver.clone(), "data"),
            binary(BinOp::Add, member(receiver, "data"), text.clone()),
        ),
        Statement::new(StmtKind::Return(Some(Expression::new(ExprKind::Tuple(
            vec![call("len", vec![text]), Expression::null()],
        ))))),
    ])
}

fn buffer_write(receiver: Expression, args: &[Argument]) -> Expression {
    let data = arg_or_null(args, 0);
    buffer_write_string(
        receiver,
        &[Argument::positional(bytes_to_string(data.clone()))],
    )
}

fn buffer_write_byte(receiver: Expression, args: &[Argument]) -> Expression {
    let text = call("__go_str_from_char_code", vec![arg_or_null(args, 0)]);
    lambda_call(vec![
        assign_stmt(
            member(receiver.clone(), "data"),
            binary(BinOp::Add, member(receiver, "data"), text),
        ),
        Statement::new(StmtKind::Return(Some(Expression::null()))),
    ])
}

fn buffer_write_rune(receiver: Expression, args: &[Argument]) -> Expression {
    let text = Expression::ident("__go_buffer_rune_text");
    lambda_call(vec![
        var_decl(
            "__go_buffer_rune_text",
            call("__go_str_from_char_code", vec![arg_or_null(args, 0)]),
        ),
        assign_stmt(
            member(receiver.clone(), "data"),
            binary(BinOp::Add, member(receiver, "data"), text.clone()),
        ),
        Statement::new(StmtKind::Return(Some(Expression::new(ExprKind::Tuple(
            vec![call("len", vec![text]), Expression::null()],
        ))))),
    ])
}

pub(crate) fn buffer_string(receiver: Expression) -> Expression {
    string_slice(
        member(receiver.clone(), "data"),
        member(receiver.clone(), "pos"),
        call("len", vec![member(receiver, "data")]),
    )
}

fn buffer_len(receiver: Expression) -> Expression {
    binary(
        BinOp::Sub,
        call("len", vec![member(receiver.clone(), "data")]),
        member(receiver, "pos"),
    )
}

fn buffer_reset(receiver: Expression) -> Expression {
    lambda_call(vec![
        assign_stmt(member(receiver.clone(), "data"), Expression::string("")),
        assign_stmt(member(receiver, "pos"), Expression::int(0)),
        Statement::new(StmtKind::Return(Some(Expression::null()))),
    ])
}

fn buffer_grow(receiver: Expression, args: &[Argument]) -> Expression {
    let desired = binary(
        BinOp::Add,
        call("len", vec![member(receiver.clone(), "data")]),
        arg_or_null(args, 0),
    );
    lambda_call(vec![
        assign_stmt(member(receiver, "gob_len"), desired),
        Statement::new(StmtKind::Return(Some(Expression::null()))),
    ])
}

fn buffer_cap(receiver: Expression) -> Expression {
    let len = call("len", vec![member(receiver.clone(), "data")]);
    let cap = member(receiver, "gob_len");
    Expression::new(ExprKind::Ternary {
        cond: Box::new(binary(BinOp::Gt, cap.clone(), len.clone())),
        then: Box::new(cap),
        else_: Box::new(len),
    })
}

fn buffer_available(receiver: Expression) -> Expression {
    let cap = buffer_cap(receiver.clone());
    let len = call("len", vec![member(receiver, "data")]);
    let available = binary(BinOp::Sub, cap.clone(), len);
    Expression::new(ExprKind::Ternary {
        cond: Box::new(binary(BinOp::Gt, available.clone(), Expression::int(0))),
        then: Box::new(available),
        else_: Box::new(Expression::int(0)),
    })
}

fn buffer_next(receiver: Expression, args: &[Argument]) -> Expression {
    let n = arg_or_null(args, 0);
    let chunk = Expression::ident("__go_buffer_chunk");
    Expression::new(ExprKind::Cast {
        expr: Box::new(lambda_call(vec![
            var_decl(
                "__go_buffer_chunk",
                string_slice(
                    member(receiver.clone(), "data"),
                    member(receiver.clone(), "pos"),
                    binary(BinOp::Add, member(receiver.clone(), "pos"), n),
                ),
            ),
            assign_stmt(
                member(receiver.clone(), "pos"),
                binary(
                    BinOp::Add,
                    member(receiver.clone(), "pos"),
                    call("len", vec![chunk.clone()]),
                ),
            ),
            Statement::new(StmtKind::Return(Some(call(
                "__go_io_string_to_bytes",
                vec![chunk],
            )))),
        ])),
        type_name: "[]byte".to_string(),
    })
}

fn buffer_truncate(receiver: Expression, args: &[Argument]) -> Expression {
    let n = arg_or_null(args, 0);
    lambda_call(vec![
        assign_stmt(
            member(receiver.clone(), "data"),
            string_slice(
                member(receiver.clone(), "data"),
                Expression::int(0),
                binary(BinOp::Add, member(receiver, "pos"), n),
            ),
        ),
        Statement::new(StmtKind::Return(Some(Expression::null()))),
    ])
}

fn buffer_read_rune(receiver: Expression) -> Expression {
    let chunk = Expression::ident("__go_buffer_rune");
    lambda_call(vec![
        var_decl(
            "__go_buffer_rune",
            string_slice(
                member(receiver.clone(), "data"),
                member(receiver.clone(), "pos"),
                binary(
                    BinOp::Add,
                    member(receiver.clone(), "pos"),
                    Expression::int(1),
                ),
            ),
        ),
        assign_stmt(
            member(receiver.clone(), "pos"),
            binary(
                BinOp::Add,
                member(receiver, "pos"),
                call("len", vec![chunk.clone()]),
            ),
        ),
        Statement::new(StmtKind::Return(Some(Expression::new(ExprKind::Tuple(
            vec![chunk.clone(), call("len", vec![chunk]), Expression::null()],
        ))))),
    ])
}

fn buffer_unread_rune(receiver: Expression) -> Expression {
    let previous = binary(
        BinOp::Sub,
        member(receiver.clone(), "pos"),
        Expression::int(1),
    );
    lambda_call(vec![
        assign_stmt(
            member(receiver, "pos"),
            Expression::new(ExprKind::Ternary {
                cond: Box::new(binary(BinOp::Gt, previous.clone(), Expression::int(0))),
                then: Box::new(previous),
                else_: Box::new(Expression::int(0)),
            }),
        ),
        Statement::new(StmtKind::Return(Some(Expression::null()))),
    ])
}

fn buffer_read_from(receiver: Expression, args: &[Argument]) -> Expression {
    let text = Expression::ident("__go_buffer_text");
    let receiver_value = value_arg(receiver);
    lambda_call(vec![
        var_decl(
            "__go_buffer_text",
            reader_all(value_arg(arg_or_null(args, 0))),
        ),
        expr_stmt(buffer_write_string(
            receiver_value,
            &[Argument::positional(text.clone())],
        )),
        Statement::new(StmtKind::Return(Some(Expression::new(ExprKind::Tuple(
            vec![call("len", vec![text]), Expression::null()],
        ))))),
    ])
}

fn buffer_write_to(
    pointer_receiver: Expression,
    value_receiver: Expression,
    args: &[Argument],
) -> Expression {
    let text = Expression::ident("__go_buffer_text");
    let target = mutable_value_arg(arg_or_null(args, 0));
    lambda_call(vec![
        var_decl(
            "__go_buffer_text",
            buffer_string(value_arg(pointer_receiver)),
        ),
        expr_stmt(Expression::new(ExprKind::Assign {
            target: Box::new(member(target.clone(), "data")),
            value: Box::new(binary(BinOp::Add, member(target, "data"), text.clone())),
        })),
        expr_stmt(Expression::new(ExprKind::Assign {
            target: Box::new(member(value_receiver.clone(), "pos")),
            value: Box::new(call("len", vec![member(value_receiver, "data")])),
        })),
        Statement::new(StmtKind::Return(Some(Expression::new(ExprKind::Tuple(
            vec![call("len", vec![text]), Expression::null()],
        ))))),
    ])
}

fn buffer_equal(receiver: Expression, args: &[Argument]) -> Expression {
    binary(
        BinOp::Eq,
        buffer_string(value_arg(receiver)),
        buffer_string(value_arg(pointer_arg(arg_or_null(args, 0)))),
    )
}

fn io_read_all(args: &[Argument]) -> Expression {
    Expression::new(ExprKind::Tuple(vec![
        string_to_bytes(reader_all(value_arg(arg_or_null(args, 0)))),
        Expression::null(),
    ]))
}

fn io_limit_reader(args: &[Argument]) -> Expression {
    let reader = value_arg(arg_or_null(args, 0));
    let n = arg_or_null(args, 1);
    pointer_arg(reader_object(string_slice(
        reader_text(reader),
        Expression::int(0),
        n,
    )))
}

fn io_multi_reader(args: &[Argument]) -> Expression {
    let mut parts = arg_values(args)
        .into_iter()
        .map(value_arg)
        .map(reader_all)
        .collect::<Vec<_>>();
    let data = if parts.is_empty() {
        Expression::string("")
    } else {
        let first = parts.remove(0);
        parts
            .into_iter()
            .fold(first, |acc, item| binary(BinOp::Add, acc, item))
    };
    pointer_arg(reader_object(data))
}

fn io_tee_reader(args: &[Argument]) -> Expression {
    pointer_arg(typed_object(
        "__goReader",
        vec![
            ("data", reader_text(value_arg(arg_or_null(args, 0)))),
            ("pos", Expression::int(0)),
            ("last", Expression::int(0)),
            ("tee", value_arg(arg_or_null(args, 1))),
        ],
    ))
}

fn io_write_string(args: &[Argument]) -> Expression {
    buffer_write_string(value_arg(arg_or_null(args, 0)), &[arg_or_null_arg(args, 1)])
}

fn io_copy(args: &[Argument], limit: Option<Expression>) -> Expression {
    let text = Expression::ident("__go_io_copy_text");
    let source = value_arg(arg_or_null(args, 1));
    let read_expr = match limit {
        Some(n) => reader_take(source, n),
        None => reader_all(source),
    };
    lambda_call(vec![
        var_decl("__go_io_copy_text", read_expr),
        expr_stmt(buffer_write_string(
            value_arg(arg_or_null(args, 0)),
            &[Argument::positional(text.clone())],
        )),
        Statement::new(StmtKind::Return(Some(Expression::new(ExprKind::Tuple(
            vec![call("len", vec![text]), Expression::null()],
        ))))),
    ])
}

fn io_read_at_least(args: &[Argument], min: Expression) -> Expression {
    reader_read_into_with_min(value_arg(arg_or_null(args, 0)), arg_or_null(args, 1), min)
}

pub(crate) fn rewrite_io_method_call(
    object: Expression,
    receiver_type: &str,
    field: &str,
    args: &[Argument],
) -> Option<Expression> {
    match receiver_type.trim().trim_start_matches('*') {
        "__goReader" => rewrite_reader_method_call(object, receiver_type, field, args),
        "__goScanner" => rewrite_scanner_method_call(object, receiver_type, field, args),
        "__goBufioWriter" => rewrite_writer_method_call(object, receiver_type, field, args),
        _ => None,
    }
}

fn rewrite_reader_method_call(
    object: Expression,
    receiver_type: &str,
    field: &str,
    args: &[Argument],
) -> Option<Expression> {
    let receiver = receiver_value(object, receiver_type);
    match field {
        "Len" | "Buffered" => Some(call("len", vec![reader_text(receiver)])),
        "Size" => Some(Expression::new(ExprKind::Ternary {
            cond: Box::new(binary(BinOp::Eq, receiver.clone(), Expression::null())),
            then: Box::new(Expression::int(0)),
            else_: Box::new(call("len", vec![member(receiver, "data")])),
        })),
        "ReadByte" => Some(reader_read_byte(receiver)),
        "UnreadByte" | "UnreadRune" => Some(reader_unread(receiver)),
        "ReadRune" => Some(reader_read_rune(receiver)),
        "ReadAt" => Some(reader_read_at(receiver, args)),
        "Peek" => Some(Expression::new(ExprKind::Tuple(vec![
            string_to_bytes(string_slice(
                reader_text(receiver),
                Expression::int(0),
                arg_or_null(args, 0),
            )),
            Expression::null(),
        ]))),
        "ReadSlice" | "ReadBytes" => Some(reader_read_delim(receiver, args, true)),
        "ReadString" => Some(reader_read_delim(receiver, args, false)),
        "ReadLine" => Some(reader_read_line(receiver)),
        "Discard" => Some(reader_discard(receiver, args)),
        "Close" => Some(Expression::null()),
        "Seek" => Some(reader_seek(receiver, args)),
        "Read" => Some(reader_read_into(receiver, args)),
        _ => None,
    }
}

fn rewrite_scanner_method_call(
    object: Expression,
    receiver_type: &str,
    field: &str,
    args: &[Argument],
) -> Option<Expression> {
    if receiver_type.trim().trim_start_matches('*') != "__goScanner" {
        return None;
    }
    let receiver = receiver_value(object, receiver_type);
    match field {
        "Split" => return Some(scanner_split(receiver, args)),
        "Scan" => return Some(scanner_scan(receiver)),
        "Text" => return Some(member(receiver, "cur")),
        "Bytes" => return Some(string_to_bytes(member(receiver, "cur"))),
        _ => return None,
    }
}

fn rewrite_writer_method_call(
    object: Expression,
    receiver_type: &str,
    field: &str,
    args: &[Argument],
) -> Option<Expression> {
    let receiver = receiver_value(object, receiver_type);
    match field {
        "WriteString" => Some(writer_write_string(receiver, args)),
        "WriteByte" => Some(writer_write_byte(receiver, args)),
        "WriteRune" => Some(writer_write_rune(receiver, args)),
        "Buffered" => Some(call("len", vec![member(receiver, "buf")])),
        "Flush" => Some(writer_flush(receiver)),
        "Reset" => Some(writer_reset(receiver, args)),
        _ => None,
    }
}

fn receiver_value(object: Expression, receiver_type: &str) -> Expression {
    if receiver_type.trim().starts_with('*') {
        Expression::new(ExprKind::RefLoad(Box::new(object)))
    } else {
        object
    }
}

pub(crate) fn reader_text(receiver: Expression) -> Expression {
    Expression::new(ExprKind::Ternary {
        cond: Box::new(binary(
            BinOp::Or,
            binary(BinOp::Eq, receiver.clone(), Expression::null()),
            binary(
                BinOp::GtEq,
                member(receiver.clone(), "pos"),
                call("len", vec![member(receiver.clone(), "data")]),
            ),
        )),
        then: Box::new(Expression::string("")),
        else_: Box::new(string_slice(
            member(receiver.clone(), "data"),
            member(receiver.clone(), "pos"),
            call("len", vec![member(receiver, "data")]),
        )),
    })
}

pub(crate) fn reader_take(receiver: Expression, n: Expression) -> Expression {
    lambda_call(vec![
        var_decl("__go_reader_obj", receiver),
        var_decl("__go_reader_n", n),
        Statement::new(StmtKind::If {
            cond: binary(
                BinOp::Or,
                binary(
                    BinOp::LtEq,
                    Expression::ident("__go_reader_n"),
                    Expression::int(0),
                ),
                binary(
                    BinOp::GtEq,
                    member(Expression::ident("__go_reader_obj"), "pos"),
                    call(
                        "len",
                        vec![member(Expression::ident("__go_reader_obj"), "data")],
                    ),
                ),
            ),
            then_body: vec![Statement::new(StmtKind::Return(Some(Expression::string(
                "",
            ))))],
            elifs: Vec::new(),
            else_body: None,
        }),
        var_decl(
            "__go_reader_remaining",
            binary(
                BinOp::Sub,
                call(
                    "len",
                    vec![member(Expression::ident("__go_reader_obj"), "data")],
                ),
                member(Expression::ident("__go_reader_obj"), "pos"),
            ),
        ),
        Statement::new(StmtKind::If {
            cond: binary(
                BinOp::Gt,
                Expression::ident("__go_reader_n"),
                Expression::ident("__go_reader_remaining"),
            ),
            then_body: vec![expr_stmt(Expression::new(ExprKind::Assign {
                target: Box::new(Expression::ident("__go_reader_n")),
                value: Box::new(Expression::ident("__go_reader_remaining")),
            }))],
            elifs: Vec::new(),
            else_body: None,
        }),
        var_decl(
            "__go_reader_start",
            member(Expression::ident("__go_reader_obj"), "pos"),
        ),
        expr_stmt(Expression::new(ExprKind::Assign {
            target: Box::new(member(Expression::ident("__go_reader_obj"), "pos")),
            value: Box::new(binary(
                BinOp::Add,
                member(Expression::ident("__go_reader_obj"), "pos"),
                Expression::ident("__go_reader_n"),
            )),
        })),
        expr_stmt(Expression::new(ExprKind::Assign {
            target: Box::new(member(Expression::ident("__go_reader_obj"), "last")),
            value: Box::new(Expression::ident("__go_reader_n")),
        })),
        var_decl(
            "__go_reader_out",
            string_slice(
                member(Expression::ident("__go_reader_obj"), "data"),
                Expression::ident("__go_reader_start"),
                member(Expression::ident("__go_reader_obj"), "pos"),
            ),
        ),
        Statement::new(StmtKind::If {
            cond: binary(
                BinOp::NotEq,
                member(Expression::ident("__go_reader_obj"), "tee"),
                Expression::null(),
            ),
            then_body: vec![expr_stmt(buffer_write_string(
                member(Expression::ident("__go_reader_obj"), "tee"),
                &[Argument::positional(Expression::ident("__go_reader_out"))],
            ))],
            elifs: Vec::new(),
            else_body: None,
        }),
        Statement::new(StmtKind::Return(Some(Expression::ident("__go_reader_out")))),
    ])
}

pub(crate) fn reader_all(receiver: Expression) -> Expression {
    reader_take(receiver.clone(), call("len", vec![reader_text(receiver)]))
}

fn reader_read_byte(receiver: Expression) -> Expression {
    lambda_call(vec![
        var_decl(
            "__go_reader_byte",
            reader_take(receiver, Expression::int(1)),
        ),
        Statement::new(StmtKind::If {
            cond: binary(
                BinOp::Eq,
                call("len", vec![Expression::ident("__go_reader_byte")]),
                Expression::int(0),
            ),
            then_body: vec![Statement::new(StmtKind::Return(Some(Expression::new(
                ExprKind::Tuple(vec![Expression::string(""), Expression::string("EOF")]),
            ))))],
            elifs: Vec::new(),
            else_body: None,
        }),
        Statement::new(StmtKind::Return(Some(Expression::new(ExprKind::Tuple(
            vec![Expression::ident("__go_reader_byte"), Expression::null()],
        ))))),
    ])
}

fn reader_read_rune(receiver: Expression) -> Expression {
    lambda_call(vec![
        var_decl(
            "__go_reader_rune",
            reader_take(receiver, Expression::int(1)),
        ),
        Statement::new(StmtKind::If {
            cond: binary(
                BinOp::Eq,
                call("len", vec![Expression::ident("__go_reader_rune")]),
                Expression::int(0),
            ),
            then_body: vec![Statement::new(StmtKind::Return(Some(Expression::new(
                ExprKind::Tuple(vec![
                    Expression::string(""),
                    Expression::int(0),
                    Expression::string("EOF"),
                ]),
            ))))],
            elifs: Vec::new(),
            else_body: None,
        }),
        Statement::new(StmtKind::Return(Some(Expression::new(ExprKind::Tuple(
            vec![
                Expression::ident("__go_reader_rune"),
                call("len", vec![Expression::ident("__go_reader_rune")]),
                Expression::null(),
            ],
        ))))),
    ])
}

fn reader_unread(receiver: Expression) -> Expression {
    lambda_call(vec![
        Statement::new(StmtKind::If {
            cond: binary(
                BinOp::Gt,
                member(receiver.clone(), "last"),
                Expression::int(0),
            ),
            then_body: vec![
                expr_stmt(Expression::new(ExprKind::Assign {
                    target: Box::new(member(receiver.clone(), "pos")),
                    value: Box::new(binary(
                        BinOp::Sub,
                        member(receiver.clone(), "pos"),
                        member(receiver.clone(), "last"),
                    )),
                })),
                Statement::new(StmtKind::If {
                    cond: binary(
                        BinOp::Lt,
                        member(receiver.clone(), "pos"),
                        Expression::int(0),
                    ),
                    then_body: vec![expr_stmt(Expression::new(ExprKind::Assign {
                        target: Box::new(member(receiver.clone(), "pos")),
                        value: Box::new(Expression::int(0)),
                    }))],
                    elifs: Vec::new(),
                    else_body: None,
                }),
                expr_stmt(Expression::new(ExprKind::Assign {
                    target: Box::new(member(receiver.clone(), "last")),
                    value: Box::new(Expression::int(0)),
                })),
            ],
            elifs: Vec::new(),
            else_body: None,
        }),
        Statement::new(StmtKind::Return(Some(Expression::null()))),
    ])
}

fn reader_read_delim(receiver: Expression, args: &[Argument], as_bytes: bool) -> Expression {
    let delim = call("__go_str_from_char_code", vec![arg_or_null(args, 0)]);
    lambda_call(vec![
        var_decl("__go_reader_text", reader_text(receiver.clone())),
        var_decl(
            "__go_reader_idx",
            call(
                "strings.Index",
                vec![Expression::ident("__go_reader_text"), delim],
            ),
        ),
        Statement::new(StmtKind::If {
            cond: binary(
                BinOp::GtEq,
                Expression::ident("__go_reader_idx"),
                Expression::int(0),
            ),
            then_body: vec![Statement::new(StmtKind::Return(Some(Expression::new(
                ExprKind::Tuple(vec![
                    maybe_bytes(
                        reader_take(
                            receiver.clone(),
                            binary(
                                BinOp::Add,
                                Expression::ident("__go_reader_idx"),
                                Expression::int(1),
                            ),
                        ),
                        as_bytes,
                    ),
                    Expression::null(),
                ]),
            ))))],
            elifs: Vec::new(),
            else_body: None,
        }),
        Statement::new(StmtKind::Return(Some(Expression::new(ExprKind::Tuple(
            vec![
                maybe_bytes(reader_all(receiver), as_bytes),
                Expression::string("EOF"),
            ],
        ))))),
    ])
}

fn maybe_bytes(expr: Expression, as_bytes: bool) -> Expression {
    if as_bytes {
        string_to_bytes(expr)
    } else {
        expr
    }
}

fn reader_read_line(receiver: Expression) -> Expression {
    lambda_call(vec![
        var_decl(
            "__go_reader_line_tuple",
            reader_read_delim(
                receiver,
                &[Argument::positional(Expression::int(10))],
                false,
            ),
        ),
        Statement::new(StmtKind::Return(Some(Expression::new(ExprKind::Tuple(
            vec![
                string_to_bytes(call(
                    "go.strings.TrimSuffix",
                    vec![
                        Expression::new(ExprKind::Index {
                            object: Box::new(Expression::ident("__go_reader_line_tuple")),
                            index: Box::new(Expression::int(0)),
                            null_safe: false,
                        }),
                        Expression::string("\n"),
                    ],
                )),
                Expression::bool(false),
                Expression::new(ExprKind::Index {
                    object: Box::new(Expression::ident("__go_reader_line_tuple")),
                    index: Box::new(Expression::int(1)),
                    null_safe: false,
                }),
            ],
        ))))),
    ])
}

fn reader_discard(receiver: Expression, args: &[Argument]) -> Expression {
    lambda_call(vec![
        var_decl(
            "__go_reader_discard",
            reader_take(receiver, arg_or_null(args, 0)),
        ),
        Statement::new(StmtKind::Return(Some(Expression::new(ExprKind::Tuple(
            vec![
                call("len", vec![Expression::ident("__go_reader_discard")]),
                Expression::null(),
            ],
        ))))),
    ])
}

fn reader_seek(receiver: Expression, args: &[Argument]) -> Expression {
    lambda_call(vec![
        var_decl("__go_reader_next", arg_or_null(args, 0)),
        Statement::new(StmtKind::If {
            cond: binary(BinOp::Eq, arg_or_null(args, 1), Expression::int(1)),
            then_body: vec![expr_stmt(Expression::new(ExprKind::Assign {
                target: Box::new(Expression::ident("__go_reader_next")),
                value: Box::new(binary(
                    BinOp::Add,
                    member(receiver.clone(), "pos"),
                    arg_or_null(args, 0),
                )),
            }))],
            elifs: Vec::new(),
            else_body: Some(vec![Statement::new(StmtKind::If {
                cond: binary(BinOp::Eq, arg_or_null(args, 1), Expression::int(2)),
                then_body: vec![expr_stmt(Expression::new(ExprKind::Assign {
                    target: Box::new(Expression::ident("__go_reader_next")),
                    value: Box::new(binary(
                        BinOp::Add,
                        call("len", vec![member(receiver.clone(), "data")]),
                        arg_or_null(args, 0),
                    )),
                }))],
                elifs: Vec::new(),
                else_body: None,
            })]),
        }),
        Statement::new(StmtKind::If {
            cond: binary(
                BinOp::Lt,
                Expression::ident("__go_reader_next"),
                Expression::int(0),
            ),
            then_body: vec![expr_stmt(Expression::new(ExprKind::Assign {
                target: Box::new(Expression::ident("__go_reader_next")),
                value: Box::new(Expression::int(0)),
            }))],
            elifs: Vec::new(),
            else_body: None,
        }),
        Statement::new(StmtKind::If {
            cond: binary(
                BinOp::Gt,
                Expression::ident("__go_reader_next"),
                call("len", vec![member(receiver.clone(), "data")]),
            ),
            then_body: vec![expr_stmt(Expression::new(ExprKind::Assign {
                target: Box::new(Expression::ident("__go_reader_next")),
                value: Box::new(call("len", vec![member(receiver.clone(), "data")])),
            }))],
            elifs: Vec::new(),
            else_body: None,
        }),
        expr_stmt(Expression::new(ExprKind::Assign {
            target: Box::new(member(receiver.clone(), "pos")),
            value: Box::new(Expression::ident("__go_reader_next")),
        })),
        Statement::new(StmtKind::Return(Some(Expression::new(ExprKind::Tuple(
            vec![member(receiver, "pos"), Expression::null()],
        ))))),
    ])
}

fn reader_read_into(receiver: Expression, args: &[Argument]) -> Expression {
    reader_read_into_with_min(receiver, arg_or_null(args, 0), Expression::int(1))
}

fn reader_read_into_with_min(
    receiver: Expression,
    buffer: Expression,
    min: Expression,
) -> Expression {
    let buffer_target = mutable_storage_expr(&buffer);
    let mut body = vec![
        var_decl("__go_reader_buf", buffer),
        var_decl(
            "__go_reader_read_text",
            reader_take(
                receiver,
                call("len", vec![Expression::ident("__go_reader_buf")]),
            ),
        ),
        var_decl(
            "__go_reader_read_bytes",
            string_to_bytes(Expression::ident("__go_reader_read_text")),
        ),
        var_decl("__go_reader_i", Expression::int(0)),
        Statement::new(StmtKind::For {
            init: None,
            cond: Some(binary(
                BinOp::Lt,
                Expression::ident("__go_reader_i"),
                call("len", vec![Expression::ident("__go_reader_read_bytes")]),
            )),
            update: Some(Expression::new(ExprKind::Assign {
                target: Box::new(Expression::ident("__go_reader_i")),
                value: Box::new(binary(
                    BinOp::Add,
                    Expression::ident("__go_reader_i"),
                    Expression::int(1),
                )),
            })),
            body: vec![expr_stmt(Expression::new(ExprKind::Assign {
                target: Box::new(index(
                    Expression::ident("__go_reader_buf"),
                    Expression::ident("__go_reader_i"),
                )),
                value: Box::new(index(
                    Expression::ident("__go_reader_read_bytes"),
                    Expression::ident("__go_reader_i"),
                )),
            }))],
        }),
    ];
    if let Some(target) = buffer_target {
        body.push(expr_stmt(Expression::new(ExprKind::Assign {
            target: Box::new(target),
            value: Box::new(Expression::ident("__go_reader_buf")),
        })));
    }
    body.extend([
        var_decl(
            "__go_reader_n",
            call("len", vec![Expression::ident("__go_reader_read_text")]),
        ),
        Statement::new(StmtKind::Return(Some(Expression::new(ExprKind::Tuple(
            vec![
                Expression::ident("__go_reader_n"),
                Expression::new(ExprKind::Ternary {
                    cond: Box::new(binary(BinOp::Lt, Expression::ident("__go_reader_n"), min)),
                    then: Box::new(Expression::string("EOF")),
                    else_: Box::new(Expression::null()),
                }),
            ],
        ))))),
    ]);
    lambda_call(body)
}

fn reader_read_at(receiver: Expression, args: &[Argument]) -> Expression {
    lambda_call(vec![
        var_decl("__go_reader_at_buf", arg_or_null(args, 0)),
        var_decl("__go_reader_at_off", arg_or_null(args, 1)),
        var_decl(
            "__go_reader_at_text",
            string_slice(
                member(receiver, "data"),
                Expression::ident("__go_reader_at_off"),
                binary(
                    BinOp::Add,
                    Expression::ident("__go_reader_at_off"),
                    call("len", vec![Expression::ident("__go_reader_at_buf")]),
                ),
            ),
        ),
        var_decl(
            "__go_reader_at_bytes",
            string_to_bytes(Expression::ident("__go_reader_at_text")),
        ),
        var_decl("__go_reader_at_i", Expression::int(0)),
        Statement::new(StmtKind::For {
            init: None,
            cond: Some(binary(
                BinOp::Lt,
                Expression::ident("__go_reader_at_i"),
                call("len", vec![Expression::ident("__go_reader_at_bytes")]),
            )),
            update: Some(Expression::new(ExprKind::Assign {
                target: Box::new(Expression::ident("__go_reader_at_i")),
                value: Box::new(binary(
                    BinOp::Add,
                    Expression::ident("__go_reader_at_i"),
                    Expression::int(1),
                )),
            })),
            body: vec![expr_stmt(Expression::new(ExprKind::Assign {
                target: Box::new(index(
                    Expression::ident("__go_reader_at_buf"),
                    Expression::ident("__go_reader_at_i"),
                )),
                value: Box::new(index(
                    Expression::ident("__go_reader_at_bytes"),
                    Expression::ident("__go_reader_at_i"),
                )),
            }))],
        }),
        Statement::new(StmtKind::Return(Some(Expression::new(ExprKind::Tuple(
            vec![
                call("len", vec![Expression::ident("__go_reader_at_text")]),
                Expression::new(ExprKind::Ternary {
                    cond: Box::new(binary(
                        BinOp::Lt,
                        call("len", vec![Expression::ident("__go_reader_at_text")]),
                        call("len", vec![Expression::ident("__go_reader_at_buf")]),
                    )),
                    then: Box::new(Expression::string("EOF")),
                    else_: Box::new(Expression::null()),
                }),
            ],
        ))))),
    ])
}

fn scanner_split(receiver: Expression, args: &[Argument]) -> Expression {
    lambda_call(vec![
        expr_stmt(Expression::new(ExprKind::Assign {
            target: Box::new(member(receiver.clone(), "mode")),
            value: Box::new(arg_or_null(args, 0)),
        })),
        expr_stmt(Expression::new(ExprKind::Assign {
            target: Box::new(member(receiver.clone(), "pos")),
            value: Box::new(Expression::int(0)),
        })),
        Statement::new(StmtKind::Return(Some(Expression::null()))),
    ])
}

fn scanner_scan(receiver: Expression) -> Expression {
    lambda_call(vec![
        var_decl(
            "__go_scanner_tokens",
            scanner_tokens(
                member(receiver.clone(), "source"),
                member(receiver.clone(), "mode"),
            ),
        ),
        Statement::new(StmtKind::If {
            cond: binary(
                BinOp::GtEq,
                member(receiver.clone(), "pos"),
                call("len", vec![Expression::ident("__go_scanner_tokens")]),
            ),
            then_body: vec![
                expr_stmt(Expression::new(ExprKind::Assign {
                    target: Box::new(member(receiver.clone(), "cur")),
                    value: Box::new(Expression::string("")),
                })),
                Statement::new(StmtKind::Return(Some(Expression::bool(false)))),
            ],
            elifs: Vec::new(),
            else_body: None,
        }),
        expr_stmt(Expression::new(ExprKind::Assign {
            target: Box::new(member(receiver.clone(), "cur")),
            value: Box::new(Expression::new(ExprKind::Index {
                object: Box::new(Expression::ident("__go_scanner_tokens")),
                index: Box::new(member(receiver.clone(), "pos")),
                null_safe: false,
            })),
        })),
        expr_stmt(Expression::new(ExprKind::Assign {
            target: Box::new(member(receiver.clone(), "pos")),
            value: Box::new(binary(
                BinOp::Add,
                member(receiver, "pos"),
                Expression::int(1),
            )),
        })),
        Statement::new(StmtKind::Return(Some(Expression::bool(true)))),
    ])
}

fn scanner_tokens(source: Expression, mode: Expression) -> Expression {
    Expression::new(ExprKind::Ternary {
        cond: Box::new(binary(
            BinOp::Eq,
            mode.clone(),
            Expression::string("ScanWords"),
        )),
        then: Box::new(call("strings.Fields", vec![source.clone()])),
        else_: Box::new(Expression::new(ExprKind::Ternary {
            cond: Box::new(binary(BinOp::Eq, mode, Expression::string("ScanBytes"))),
            then: Box::new(call("ecma.array.from", vec![source.clone()])),
            else_: Box::new(call(
                "strings.Split",
                vec![source, Expression::string("\n")],
            )),
        })),
    })
}

fn writer_write_string(receiver: Expression, args: &[Argument]) -> Expression {
    let text = arg_or_null(args, 0);
    lambda_call(vec![
        expr_stmt(Expression::new(ExprKind::Assign {
            target: Box::new(member(receiver.clone(), "buf")),
            value: Box::new(binary(BinOp::Add, member(receiver, "buf"), text.clone())),
        })),
        Statement::new(StmtKind::Return(Some(Expression::new(ExprKind::Tuple(
            vec![call("len", vec![text]), Expression::null()],
        ))))),
    ])
}

fn writer_write_byte(receiver: Expression, args: &[Argument]) -> Expression {
    lambda_call(vec![
        expr_stmt(Expression::new(ExprKind::Assign {
            target: Box::new(member(receiver.clone(), "buf")),
            value: Box::new(binary(
                BinOp::Add,
                member(receiver, "buf"),
                call("__go_str_from_char_code", vec![arg_or_null(args, 0)]),
            )),
        })),
        Statement::new(StmtKind::Return(Some(Expression::null()))),
    ])
}

fn writer_write_rune(receiver: Expression, args: &[Argument]) -> Expression {
    writer_write_string(
        receiver,
        &[Argument::positional(call(
            "__go_str_from_char_code",
            vec![arg_or_null(args, 0)],
        ))],
    )
}

fn writer_flush(receiver: Expression) -> Expression {
    lambda_call(vec![
        Statement::new(StmtKind::If {
            cond: binary(
                BinOp::NotEq,
                member(receiver.clone(), "out"),
                Expression::null(),
            ),
            then_body: vec![expr_stmt(buffer_write_string(
                member(receiver.clone(), "out"),
                &[Argument::positional(member(receiver.clone(), "buf"))],
            ))],
            elifs: Vec::new(),
            else_body: None,
        }),
        expr_stmt(Expression::new(ExprKind::Assign {
            target: Box::new(member(receiver, "buf")),
            value: Box::new(Expression::string("")),
        })),
        Statement::new(StmtKind::Return(Some(Expression::null()))),
    ])
}

fn writer_reset(receiver: Expression, args: &[Argument]) -> Expression {
    lambda_call(vec![
        expr_stmt(Expression::new(ExprKind::Assign {
            target: Box::new(member(receiver.clone(), "buf")),
            value: Box::new(Expression::string("")),
        })),
        expr_stmt(Expression::new(ExprKind::Assign {
            target: Box::new(member(receiver, "out")),
            value: Box::new(arg_or_null(args, 0)),
        })),
        Statement::new(StmtKind::Return(Some(Expression::null()))),
    ])
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

fn buffer_object(data: Expression) -> Expression {
    typed_object(
        "__goBuffer",
        vec![
            ("data", data),
            ("pos", Expression::int(0)),
            ("last", Expression::int(0)),
            ("tee", Expression::null()),
            ("gob", Expression::new(ExprKind::Array(Vec::new()))),
            ("gob0", Expression::null()),
            ("gob1", Expression::null()),
            ("gob2", Expression::null()),
            ("gob3", Expression::null()),
            ("gob4", Expression::null()),
            ("gob5", Expression::null()),
            ("gob6", Expression::null()),
            ("gob7", Expression::null()),
            ("gob_len", Expression::int(0)),
        ],
    )
}

fn reader_object(data: Expression) -> Expression {
    typed_object(
        "__goReader",
        vec![
            ("data", data),
            ("pos", Expression::int(0)),
            ("last", Expression::int(0)),
            ("tee", Expression::null()),
        ],
    )
}

fn scanner_object(source: Expression, mode: Expression) -> Expression {
    typed_object(
        "__goScanner",
        vec![
            ("tokens", Expression::new(ExprKind::Array(Vec::new()))),
            ("pos", Expression::int(0)),
            ("cur", Expression::string("")),
            ("mode", mode),
            ("source", source),
        ],
    )
}

fn writer_object(out: Expression) -> Expression {
    typed_object(
        "__goBufioWriter",
        vec![("out", out), ("buf", Expression::string(""))],
    )
}

pub(crate) fn typed_object(type_name: &str, fields: Vec<(&str, Expression)>) -> Expression {
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

fn arg_or_null(args: &[Argument], index: usize) -> Expression {
    args.get(index)
        .map(|arg| arg.value.clone())
        .unwrap_or_else(Expression::null)
}

fn arg_or_null_arg(args: &[Argument], index: usize) -> Argument {
    Argument::positional(arg_or_null(args, index))
}

fn arg_values(args: &[Argument]) -> Vec<Expression> {
    args.iter().map(|arg| arg.value.clone()).collect()
}

pub(crate) fn pointer_arg(expr: Expression) -> Expression {
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

pub(crate) fn value_arg(expr: Expression) -> Expression {
    match expr.kind {
        ExprKind::Unary {
            op: UnaryOp::AddrOf,
            expr,
        } => *expr,
        ExprKind::RefOf(_) => Expression::new(ExprKind::RefLoad(Box::new(expr))),
        _ => expr,
    }
}

fn mutable_value_arg(expr: Expression) -> Expression {
    mutable_storage_expr(&expr).unwrap_or_else(|| value_arg(expr))
}

fn mutable_storage_expr(expr: &Expression) -> Option<Expression> {
    match &expr.kind {
        ExprKind::Unary {
            op: UnaryOp::AddrOf,
            expr,
        } => Some(expr.as_ref().clone()),
        ExprKind::RefOf(place) => Some(expr_from_place(place)),
        _ => PlaceExpr::from_expr(expr).map(|place| expr_from_place(&place)),
    }
}

fn expr_from_place(place: &PlaceExpr) -> Expression {
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

pub(crate) fn member(object: Expression, field: &str) -> Expression {
    Expression::new(ExprKind::Index {
        object: Box::new(object),
        index: Box::new(Expression::string(field)),
        null_safe: false,
    })
}

pub(crate) fn index(object: Expression, index: Expression) -> Expression {
    Expression::new(ExprKind::Index {
        object: Box::new(object),
        index: Box::new(index),
        null_safe: false,
    })
}

fn call_name(expr: &Expression) -> Option<String> {
    match &expr.kind {
        ExprKind::Ident(name) => Some(name.clone()),
        ExprKind::Member { object, field, .. } => Some(format!("{}.{}", call_name(object)?, field)),
        _ => None,
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use vybe_ast::Literal;

    fn single_call_arg(expr: &Expression, expected_name: &str) -> Option<Expression> {
        let ExprKind::Call { callee, args, .. } = &expr.kind else {
            return None;
        };
        if call_name(callee).as_deref() != Some(expected_name) || args.len() != 1 {
            return None;
        }
        Some(args[0].value.clone())
    }

    #[test]
    fn bytes_string_round_trip_unwraps_to_original_string() {
        let expr = bytes_to_string(string_to_bytes(Expression::string("xy")));
        assert!(matches!(expr.kind, ExprKind::Lit(Literal::Str(ref s)) if s == "xy"));
    }

    #[test]
    fn string_bytes_round_trip_unwraps_to_original_bytes() {
        let bytes = Expression::ident("payload");
        let expr = string_to_bytes(bytes_to_string(bytes.clone()));
        assert!(matches!(expr.kind, ExprKind::Ident(ref name) if name == "payload"));
    }

    #[test]
    fn plain_bytes_to_string_still_uses_core_adapter() {
        let expr = bytes_to_string(Expression::ident("payload"));
        assert!(single_call_arg(&expr, "__go_io_bytes_to_string").is_some());
    }
}

pub(crate) fn string_slice(object: Expression, start: Expression, end: Expression) -> Expression {
    call("__go_slices_slice_common", vec![object, start, end])
}

pub(crate) fn binary(op: BinOp, left: Expression, right: Expression) -> Expression {
    Expression::new(ExprKind::Binary {
        op,
        left: Box::new(left),
        right: Box::new(right),
    })
}

pub(crate) fn var_decl(name: &str, init: Expression) -> Statement {
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

pub(crate) fn expr_stmt(expr: Expression) -> Statement {
    Statement::new(StmtKind::Expr(expr))
}

fn assign_stmt(target: Expression, value: Expression) -> Statement {
    Statement::new(StmtKind::Assign {
        targets: vec![target],
        value,
        by_ref: false,
    })
}

pub(crate) fn lambda_call(body: Vec<Statement>) -> Expression {
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

pub(crate) fn call(name: &str, args: Vec<Expression>) -> Expression {
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

use vybe_ast::{
    Argument, ArrayElement, BinOp, BindingPattern, ExprKind, Expression, LambdaBody,
    ObjectProperty, Statement, StmtKind, VarDeclKind, VarDeclarator,
};
use vybe_compiler::primitives::namespaces::{NamespaceNode, Subtree};

pub(crate) fn register_tree(root: &mut Subtree) {
    register_type(
        root,
        "encoding.base64.Encoding",
        &[
            "EncodedLen",
            "DecodedLen",
            "WithPadding",
            "EncodeToString",
            "Decode",
            "DecodeString",
        ],
        &[
            ("EncodedLen", "int"),
            ("DecodedLen", "int"),
            ("WithPadding", "__goBase64Encoding"),
            ("EncodeToString", "string"),
            ("Decode", "tuple"),
            ("DecodeString", "tuple"),
        ],
    );
    register_type(
        root,
        "__goBase64Encoding",
        &[
            "EncodedLen",
            "DecodedLen",
            "WithPadding",
            "EncodeToString",
            "Decode",
            "DecodeString",
        ],
        &[
            ("EncodedLen", "int"),
            ("DecodedLen", "int"),
            ("WithPadding", "__goBase64Encoding"),
            ("EncodeToString", "string"),
            ("Decode", "tuple"),
            ("DecodeString", "tuple"),
        ],
    );
    register_type(
        root,
        "encoding.binary.ByteOrder",
        &[
            "PutUint16",
            "Uint16",
            "PutInt16",
            "PutUint32",
            "Uint32",
            "Int32",
            "PutUint64",
            "Uint64",
            "AppendUint16",
            "AppendUint32",
        ],
        &[
            ("Uint16", "int"),
            ("Uint32", "int"),
            ("Int32", "int"),
            ("Uint64", "int"),
            ("AppendUint16", "[]byte"),
            ("AppendUint32", "[]byte"),
        ],
    );
    register_type(
        root,
        "__goByteOrder",
        &[
            "PutUint16",
            "Uint16",
            "PutInt16",
            "PutUint32",
            "Uint32",
            "Int32",
            "PutUint64",
            "Uint64",
            "AppendUint16",
            "AppendUint32",
        ],
        &[
            ("Uint16", "int"),
            ("Uint32", "int"),
            ("Int32", "int"),
            ("Uint64", "int"),
            ("AppendUint16", "[]byte"),
            ("AppendUint32", "[]byte"),
        ],
    );
    register_type(
        root,
        "__goHexDumper",
        &["Write", "Close"],
        &[("Write", "tuple")],
    );

    for (name, emit) in [
        ("hex.EncodedLen", "go.encoding.hex.EncodedLen"),
        ("hex.DecodedLen", "go.encoding.hex.DecodedLen"),
        ("hex.Encode", "go.encoding.hex.Encode"),
        ("hex.EncodeToString", "go.encoding.hex.EncodeToString"),
        ("hex.AppendEncode", "go.encoding.hex.AppendEncode"),
        ("hex.Decode", "go.encoding.hex.Decode"),
        ("hex.DecodeString", "go.encoding.hex.DecodeString"),
        ("hex.Dump", "go.encoding.hex.Dump"),
        ("hex.Dumper", "go.encoding.hex.Dumper"),
        ("binary.PutUvarint", "go.encoding.binary.PutUvarint"),
        ("binary.Uvarint", "go.encoding.binary.Uvarint"),
        ("binary.PutVarint", "go.encoding.binary.PutVarint"),
        ("binary.Varint", "go.encoding.binary.Varint"),
        ("binary.AppendUvarint", "go.encoding.binary.AppendUvarint"),
        ("binary.Size", "go.encoding.binary.Size"),
        ("binary.Read", "go.encoding.binary.Read"),
        ("binary.Write", "go.encoding.binary.Write"),
        ("binary.ReadFull", "go.encoding.binary.ReadFull"),
    ] {
        insert_path(root, name, NamespaceNode::CommonEmit(emit.to_string()));
    }

    for name in [
        "base64.StdEncoding",
        "base64.RawStdEncoding",
        "base64.URLEncoding",
        "binary.BigEndian",
        "binary.LittleEndian",
        "binary.NativeEndian",
        "binary.MaxVarintLen64",
        "hex.InvalidByte",
    ] {
        insert_path(root, name, NamespaceNode::CommonEmit(format!("go.{name}")));
    }
}

pub(crate) fn rewrite_member(package: &str, field: &str) -> Option<Expression> {
    match (package, field) {
        ("hex", "InvalidByte") => Some(Expression::int(0)),
        ("base64", "StdEncoding") => Some(base64_encoding(false, false)),
        ("base64", "RawStdEncoding") => Some(base64_encoding(true, false)),
        ("base64", "URLEncoding") => Some(base64_encoding(false, true)),
        ("binary", "BigEndian") => Some(byte_order(false)),
        ("binary", "LittleEndian") | ("binary", "NativeEndian") => Some(byte_order(true)),
        ("binary", "MaxVarintLen64") => Some(Expression::int(10)),
        _ => None,
    }
}

pub(crate) fn rewrite_call(call_name: &str, args: &[Argument]) -> Option<Expression> {
    match call_name {
        "hex.EncodedLen" => Some(binary(BinOp::Mul, arg(args, 0), Expression::int(2))),
        "hex.DecodedLen" => Some(binary(BinOp::Div, arg(args, 0), Expression::int(2))),
        "hex.Encode" => Some(hex_encode_into(arg(args, 0), arg(args, 1))),
        "hex.EncodeToString" => Some(call("__go_hex_encode_to_string", vec![arg(args, 0)])),
        "hex.AppendEncode" => Some(call(
            "__go_array_concat",
            vec![
                arg(args, 0),
                call(
                    "__go_io_string_to_bytes",
                    vec![call("__go_hex_encode_to_string", vec![arg(args, 1)])],
                ),
            ],
        )),
        "hex.Decode" => Some(hex_decode_into(arg(args, 0), arg(args, 1))),
        "hex.DecodeString" => Some(tuple(vec![
            call("__go_hex_decode_string_bytes", vec![arg(args, 0)]),
            Expression::null(),
        ])),
        "hex.Dump" => Some(hex_dump(arg(args, 0))),
        "hex.Dumper" => Some(typed_composite(
            object(vec![("out", arg(args, 0))]),
            "__goHexDumper",
        )),
        "binary.PutUvarint" => Some(binary_put_uvarint(arg(args, 0), arg(args, 1))),
        "binary.Uvarint" => Some(binary_uvarint(arg(args, 0), false)),
        "binary.PutVarint" => Some(binary_put_varint(arg(args, 0), arg(args, 1))),
        "binary.Varint" => Some(binary_varint(arg(args, 0))),
        "binary.AppendUvarint" => Some(call(
            "__go_array_concat",
            vec![
                arg(args, 0),
                binary_append_uvarint_bytes(arg(args, 1), false),
            ],
        )),
        "binary.Size" => Some(Expression::int(2)),
        "binary.Read" => Some(Expression::null()),
        "binary.Write" => Some(Expression::null()),
        "binary.ReadFull" => Some(call_method(arg(args, 0), "Read", vec![arg(args, 1)])),
        "binary.BigEndian.PutUint16" => Some(emit_order_call(
            "__go_emit_binary_PutUint16",
            false,
            arg_values(args),
        )),
        "binary.LittleEndian.PutUint16" | "binary.NativeEndian.PutUint16" => Some(emit_order_call(
            "__go_emit_binary_PutUint16",
            true,
            arg_values(args),
        )),
        "binary.BigEndian.Uint16" => Some(emit_order_call(
            "__go_emit_binary_Uint16",
            false,
            arg_values(args),
        )),
        "binary.LittleEndian.Uint16" | "binary.NativeEndian.Uint16" => Some(emit_order_call(
            "__go_emit_binary_Uint16",
            true,
            arg_values(args),
        )),
        "binary.BigEndian.PutInt16" => Some(emit_order_call(
            "__go_emit_binary_PutInt16",
            false,
            arg_values(args),
        )),
        "binary.LittleEndian.PutInt16" | "binary.NativeEndian.PutInt16" => Some(emit_order_call(
            "__go_emit_binary_PutInt16",
            true,
            arg_values(args),
        )),
        "binary.BigEndian.PutUint32" => Some(emit_order_call(
            "__go_emit_binary_PutUint32",
            false,
            arg_values(args),
        )),
        "binary.LittleEndian.PutUint32" | "binary.NativeEndian.PutUint32" => Some(emit_order_call(
            "__go_emit_binary_PutUint32",
            true,
            arg_values(args),
        )),
        "binary.BigEndian.Uint32" => Some(emit_order_call(
            "__go_emit_binary_Uint32",
            false,
            arg_values(args),
        )),
        "binary.LittleEndian.Uint32" | "binary.NativeEndian.Uint32" => Some(emit_order_call(
            "__go_emit_binary_Uint32",
            true,
            arg_values(args),
        )),
        "binary.BigEndian.Int32" => Some(emit_order_call(
            "__go_emit_binary_Int32",
            false,
            arg_values(args),
        )),
        "binary.LittleEndian.Int32" | "binary.NativeEndian.Int32" => Some(emit_order_call(
            "__go_emit_binary_Int32",
            true,
            arg_values(args),
        )),
        "binary.BigEndian.PutUint64" => Some(put_uint64(false, arg(args, 0), arg(args, 1))),
        "binary.LittleEndian.PutUint64" | "binary.NativeEndian.PutUint64" => {
            Some(put_uint64(true, arg(args, 0), arg(args, 1)))
        }
        "binary.BigEndian.Uint64" => Some(uint64_from_bytes(false, arg(args, 0))),
        "binary.LittleEndian.Uint64" | "binary.NativeEndian.Uint64" => {
            Some(uint64_from_bytes(true, arg(args, 0)))
        }
        "binary.BigEndian.AppendUint16" => Some(emit_order_call(
            "__go_emit_binary_AppendUint16",
            false,
            arg_values(args),
        )),
        "binary.LittleEndian.AppendUint16" | "binary.NativeEndian.AppendUint16" => Some(
            emit_order_call("__go_emit_binary_AppendUint16", true, arg_values(args)),
        ),
        "binary.BigEndian.AppendUint32" => Some(emit_order_call(
            "__go_emit_binary_AppendUint32",
            false,
            arg_values(args),
        )),
        "binary.LittleEndian.AppendUint32" | "binary.NativeEndian.AppendUint32" => Some(
            emit_order_call("__go_emit_binary_AppendUint32", true, arg_values(args)),
        ),
        _ => None,
    }
}

pub(crate) fn rewrite_method_call(
    receiver: Expression,
    receiver_type: &str,
    field: &str,
    args: &[Argument],
) -> Option<Expression> {
    let ty = receiver_type.trim().trim_start_matches('*');
    match ty {
        "__goBase64Encoding" | "base64.Encoding" | "encoding.base64.Encoding" => {
            rewrite_base64_method(receiver, field, args)
        }
        "__goByteOrder" | "binary.ByteOrder" | "encoding.binary.ByteOrder" => {
            rewrite_byte_order_method(receiver, field, args)
        }
        "__goHexDumper" => rewrite_hex_dumper_method(receiver, field, args),
        _ => None,
    }
}

pub(crate) fn call_type_hint(name: &str) -> Option<&'static str> {
    match name {
        "__go_hex_encode_to_string" => Some("string"),
        "__go_hex_decode_string_bytes" => Some("[]byte"),
        "__go_emit_binary_AppendUint16" | "__go_emit_binary_AppendUint32" => Some("[]byte"),
        "__go_emit_binary_Uint16" | "__go_emit_binary_Uint32" | "__go_emit_binary_Int32" => {
            Some("int")
        }
        _ => None,
    }
}

fn rewrite_base64_method(
    receiver: Expression,
    field: &str,
    args: &[Argument],
) -> Option<Expression> {
    match field {
        "EncodedLen" => Some(base64_encoded_len(receiver, arg(args, 0))),
        "DecodedLen" => Some(binary(
            BinOp::Mul,
            binary(BinOp::Div, arg(args, 0), Expression::int(4)),
            Expression::int(3),
        )),
        "WithPadding" => Some(typed_composite(
            object(vec![
                ("raw", Expression::bool(false)),
                ("url", member(receiver, "url")),
            ]),
            "__goBase64Encoding",
        )),
        "EncodeToString" => Some(base64_encode_to_string(receiver, arg(args, 0))),
        "DecodeString" => Some(base64_decode_string(receiver, arg(args, 0))),
        "Decode" => Some(base64_decode_into(receiver, arg(args, 0), arg(args, 1))),
        _ => None,
    }
}

fn rewrite_byte_order_method(
    receiver: Expression,
    field: &str,
    args: &[Argument],
) -> Option<Expression> {
    let values = arg_values(args);
    match field {
        "PutUint16" => Some(emit_order_expr_call(
            "__go_emit_binary_PutUint16",
            receiver,
            values,
        )),
        "Uint16" => Some(emit_order_expr_call(
            "__go_emit_binary_Uint16",
            receiver,
            values,
        )),
        "PutInt16" => Some(emit_order_expr_call(
            "__go_emit_binary_PutInt16",
            receiver,
            values,
        )),
        "PutUint32" => Some(emit_order_expr_call(
            "__go_emit_binary_PutUint32",
            receiver,
            values,
        )),
        "Uint32" => Some(emit_order_expr_call(
            "__go_emit_binary_Uint32",
            receiver,
            values,
        )),
        "Int32" => Some(emit_order_expr_call(
            "__go_emit_binary_Int32",
            receiver,
            values,
        )),
        "PutUint64" => Some(put_uint64_expr(receiver, arg(args, 0), arg(args, 1))),
        "Uint64" => Some(uint64_from_bytes_expr(receiver, arg(args, 0))),
        "AppendUint16" => Some(emit_order_expr_call(
            "__go_emit_binary_AppendUint16",
            receiver,
            values,
        )),
        "AppendUint32" => Some(emit_order_expr_call(
            "__go_emit_binary_AppendUint32",
            receiver,
            values,
        )),
        _ => None,
    }
}

fn rewrite_hex_dumper_method(
    receiver: Expression,
    field: &str,
    args: &[Argument],
) -> Option<Expression> {
    match field {
        "Write" => Some(lambda_call(vec![
            var_decl("__go_hex_dumper", receiver),
            var_decl("__go_hex_payload", arg(args, 0)),
            expr_stmt(crate::adapters::bytes_io::buffer_write_string_expr(
                member(Expression::ident("__go_hex_dumper"), "out"),
                hex_dump(Expression::ident("__go_hex_payload")),
            )),
            Statement::new(StmtKind::Return(Some(tuple(vec![
                call("len", vec![Expression::ident("__go_hex_payload")]),
                Expression::null(),
            ])))),
        ])),
        "Close" => Some(Expression::null()),
        _ => None,
    }
}

fn base64_encode_to_string(receiver: Expression, src: Expression) -> Expression {
    lambda_call(vec![
        var_decl("__go_b64_enc", receiver),
        var_decl(
            "__go_b64_out",
            call(
                "__go_btoa",
                vec![call("__go_io_bytes_to_string", vec![src])],
            ),
        ),
        Statement::new(StmtKind::If {
            cond: member(Expression::ident("__go_b64_enc"), "url"),
            then_body: vec![
                assign(
                    Expression::ident("__go_b64_out"),
                    call(
                        "strings.ReplaceAll",
                        vec![
                            Expression::ident("__go_b64_out"),
                            Expression::string("+"),
                            Expression::string("-"),
                        ],
                    ),
                ),
                assign(
                    Expression::ident("__go_b64_out"),
                    call(
                        "strings.ReplaceAll",
                        vec![
                            Expression::ident("__go_b64_out"),
                            Expression::string("/"),
                            Expression::string("_"),
                        ],
                    ),
                ),
            ],
            elifs: Vec::new(),
            else_body: None,
        }),
        Statement::new(StmtKind::If {
            cond: member(Expression::ident("__go_b64_enc"), "raw"),
            then_body: vec![assign(
                Expression::ident("__go_b64_out"),
                call(
                    "go.strings.TrimRight",
                    vec![Expression::ident("__go_b64_out"), Expression::string("=")],
                ),
            )],
            elifs: Vec::new(),
            else_body: None,
        }),
        Statement::new(StmtKind::Return(Some(Expression::ident("__go_b64_out")))),
    ])
}

fn base64_decode_string(receiver: Expression, source: Expression) -> Expression {
    lambda_call(vec![
        var_decl("__go_b64_dec", receiver),
        var_decl("__go_b64_text", source),
        Statement::new(StmtKind::If {
            cond: member(Expression::ident("__go_b64_dec"), "url"),
            then_body: vec![
                assign(
                    Expression::ident("__go_b64_text"),
                    call(
                        "strings.ReplaceAll",
                        vec![
                            Expression::ident("__go_b64_text"),
                            Expression::string("-"),
                            Expression::string("+"),
                        ],
                    ),
                ),
                assign(
                    Expression::ident("__go_b64_text"),
                    call(
                        "strings.ReplaceAll",
                        vec![
                            Expression::ident("__go_b64_text"),
                            Expression::string("_"),
                            Expression::string("/"),
                        ],
                    ),
                ),
            ],
            elifs: Vec::new(),
            else_body: None,
        }),
        Statement::new(StmtKind::If {
            cond: binary(
                BinOp::Eq,
                binary(
                    BinOp::Mod,
                    call("len", vec![Expression::ident("__go_b64_text")]),
                    Expression::int(4),
                ),
                Expression::int(2),
            ),
            then_body: vec![assign(
                Expression::ident("__go_b64_text"),
                binary(
                    BinOp::Add,
                    Expression::ident("__go_b64_text"),
                    Expression::string("=="),
                ),
            )],
            elifs: vec![(
                binary(
                    BinOp::Eq,
                    binary(
                        BinOp::Mod,
                        call("len", vec![Expression::ident("__go_b64_text")]),
                        Expression::int(4),
                    ),
                    Expression::int(3),
                ),
                vec![assign(
                    Expression::ident("__go_b64_text"),
                    binary(
                        BinOp::Add,
                        Expression::ident("__go_b64_text"),
                        Expression::string("="),
                    ),
                )],
            )],
            else_body: None,
        }),
        Statement::new(StmtKind::Return(Some(tuple(vec![
            call(
                "__go_io_string_to_bytes",
                vec![call("__go_atob", vec![Expression::ident("__go_b64_text")])],
            ),
            Expression::null(),
        ])))),
    ])
}

fn base64_decode_into(receiver: Expression, dst: Expression, src: Expression) -> Expression {
    lambda_call(vec![
        var_decl("__go_b64_dst", dst),
        var_decl(
            "__go_b64_pair",
            base64_decode_string(receiver, call("__go_io_bytes_to_string", vec![src])),
        ),
        var_decl(
            "__go_b64_bytes",
            index(Expression::ident("__go_b64_pair"), Expression::int(0)),
        ),
        copy_bytes_loop("__go_b64_dst", "__go_b64_bytes"),
        Statement::new(StmtKind::Return(Some(tuple(vec![
            call("len", vec![Expression::ident("__go_b64_bytes")]),
            index(Expression::ident("__go_b64_pair"), Expression::int(1)),
        ])))),
    ])
}

fn base64_encoded_len(receiver: Expression, n: Expression) -> Expression {
    Expression::new(ExprKind::Ternary {
        cond: Box::new(member(receiver, "raw")),
        then: Box::new(binary(
            BinOp::Div,
            binary(
                BinOp::Add,
                binary(BinOp::Mul, n.clone(), Expression::int(8)),
                Expression::int(5),
            ),
            Expression::int(6),
        )),
        else_: Box::new(binary(
            BinOp::Mul,
            binary(
                BinOp::Div,
                binary(BinOp::Add, n, Expression::int(2)),
                Expression::int(3),
            ),
            Expression::int(4),
        )),
    })
}

fn hex_encode_into(dst: Expression, src: Expression) -> Expression {
    lambda_call(vec![
        var_decl("__go_hex_dst", dst),
        var_decl(
            "__go_hex_encoded",
            call(
                "__go_io_string_to_bytes",
                vec![call("__go_hex_encode_to_string", vec![src])],
            ),
        ),
        copy_bytes_loop("__go_hex_dst", "__go_hex_encoded"),
        Statement::new(StmtKind::Return(Some(call(
            "len",
            vec![Expression::ident("__go_hex_encoded")],
        )))),
    ])
}

fn hex_decode_into(dst: Expression, src: Expression) -> Expression {
    lambda_call(vec![
        var_decl("__go_hex_dst", dst),
        var_decl(
            "__go_hex_decoded",
            call(
                "__go_hex_decode_string_bytes",
                vec![call("__go_io_bytes_to_string", vec![src])],
            ),
        ),
        copy_bytes_loop("__go_hex_dst", "__go_hex_decoded"),
        Statement::new(StmtKind::Return(Some(tuple(vec![
            call("len", vec![Expression::ident("__go_hex_decoded")]),
            Expression::null(),
        ])))),
    ])
}

fn hex_dump(src: Expression) -> Expression {
    Expression::new(ExprKind::Ternary {
        cond: Box::new(binary(
            BinOp::Eq,
            call("len", vec![src.clone()]),
            Expression::int(0),
        )),
        then: Box::new(Expression::string("")),
        else_: Box::new(binary(
            BinOp::Add,
            binary(
                BinOp::Add,
                binary(
                    BinOp::Add,
                    Expression::string("00000000  "),
                    call("__go_hex_encode_to_string", vec![src.clone()]),
                ),
                Expression::string("  |"),
            ),
            binary(
                BinOp::Add,
                call("__go_io_bytes_to_string", vec![src]),
                Expression::string("|\n"),
            ),
        )),
    })
}

fn binary_put_uvarint(buf: Expression, x: Expression) -> Expression {
    lambda_call(vec![
        var_decl("__go_uvar_buf", buf),
        var_decl("__go_uvar_x", x),
        var_decl("__go_uvar_i", Expression::int(0)),
        Statement::new(StmtKind::While {
            cond: binary(
                BinOp::GtEq,
                Expression::ident("__go_uvar_x"),
                Expression::int(128),
            ),
            body: vec![
                assign(
                    index(
                        Expression::ident("__go_uvar_buf"),
                        Expression::ident("__go_uvar_i"),
                    ),
                    binary(
                        BinOp::Add,
                        binary(
                            BinOp::Mod,
                            Expression::ident("__go_uvar_x"),
                            Expression::int(128),
                        ),
                        Expression::int(128),
                    ),
                ),
                assign(
                    Expression::ident("__go_uvar_x"),
                    binary(
                        BinOp::Shr,
                        Expression::ident("__go_uvar_x"),
                        Expression::int(7),
                    ),
                ),
                assign(
                    Expression::ident("__go_uvar_i"),
                    binary(
                        BinOp::Add,
                        Expression::ident("__go_uvar_i"),
                        Expression::int(1),
                    ),
                ),
            ],
            else_body: None,
        }),
        assign(
            index(
                Expression::ident("__go_uvar_buf"),
                Expression::ident("__go_uvar_i"),
            ),
            Expression::ident("__go_uvar_x"),
        ),
        Statement::new(StmtKind::Return(Some(binary(
            BinOp::Add,
            Expression::ident("__go_uvar_i"),
            Expression::int(1),
        )))),
    ])
}

fn binary_put_varint(buf: Expression, x: Expression) -> Expression {
    let ux = Expression::new(ExprKind::Ternary {
        cond: Box::new(binary(BinOp::Lt, x.clone(), Expression::int(0))),
        then: Box::new(binary(
            BinOp::Sub,
            binary(
                BinOp::Mul,
                binary(BinOp::Mul, x.clone(), Expression::int(-1)),
                Expression::int(2),
            ),
            Expression::int(1),
        )),
        else_: Box::new(binary(BinOp::Mul, x, Expression::int(2))),
    });
    binary_put_uvarint(buf, ux)
}

fn binary_append_uvarint_bytes(x: Expression, signed: bool) -> Expression {
    lambda_call(vec![
        var_decl(
            "__go_append_buf",
            array_of((0..10).map(|_| Expression::int(0)).collect()),
        ),
        var_decl(
            "__go_append_n",
            if signed {
                binary_put_varint(Expression::ident("__go_append_buf"), x)
            } else {
                binary_put_uvarint(Expression::ident("__go_append_buf"), x)
            },
        ),
        Statement::new(StmtKind::Return(Some(call_method(
            Expression::ident("__go_append_buf"),
            "slice",
            vec![Expression::int(0), Expression::ident("__go_append_n")],
        )))),
    ])
}

fn binary_uvarint(buf: Expression, signed: bool) -> Expression {
    lambda_call(vec![
        var_decl("__go_var_buf", buf),
        var_decl("__go_var_x", Expression::int(0)),
        var_decl("__go_var_s", Expression::int(0)),
        var_decl("__go_var_i", Expression::int(0)),
        Statement::new(StmtKind::For {
            init: None,
            cond: Some(binary(
                BinOp::Lt,
                Expression::ident("__go_var_i"),
                call("len", vec![Expression::ident("__go_var_buf")]),
            )),
            update: Some(Expression::new(ExprKind::Assign {
                target: Box::new(Expression::ident("__go_var_i")),
                value: Box::new(binary(
                    BinOp::Add,
                    Expression::ident("__go_var_i"),
                    Expression::int(1),
                )),
            })),
            body: vec![
                var_decl(
                    "__go_var_b",
                    index(
                        Expression::ident("__go_var_buf"),
                        Expression::ident("__go_var_i"),
                    ),
                ),
                Statement::new(StmtKind::If {
                    cond: binary(
                        BinOp::Lt,
                        Expression::ident("__go_var_b"),
                        Expression::int(128),
                    ),
                    then_body: vec![Statement::new(StmtKind::Return(Some(tuple(vec![
                        if signed {
                            decode_signed_varint(binary(
                                BinOp::BitOr,
                                Expression::ident("__go_var_x"),
                                binary(
                                    BinOp::Shl,
                                    Expression::ident("__go_var_b"),
                                    Expression::ident("__go_var_s"),
                                ),
                            ))
                        } else {
                            binary(
                                BinOp::BitOr,
                                Expression::ident("__go_var_x"),
                                binary(
                                    BinOp::Shl,
                                    Expression::ident("__go_var_b"),
                                    Expression::ident("__go_var_s"),
                                ),
                            )
                        },
                        binary(
                            BinOp::Add,
                            Expression::ident("__go_var_i"),
                            Expression::int(1),
                        ),
                    ]))))],
                    elifs: Vec::new(),
                    else_body: None,
                }),
                assign(
                    Expression::ident("__go_var_x"),
                    binary(
                        BinOp::BitOr,
                        Expression::ident("__go_var_x"),
                        binary(
                            BinOp::Shl,
                            binary(
                                BinOp::BitAnd,
                                Expression::ident("__go_var_b"),
                                Expression::int(127),
                            ),
                            Expression::ident("__go_var_s"),
                        ),
                    ),
                ),
                assign(
                    Expression::ident("__go_var_s"),
                    binary(
                        BinOp::Add,
                        Expression::ident("__go_var_s"),
                        Expression::int(7),
                    ),
                ),
            ],
        }),
        Statement::new(StmtKind::Return(Some(tuple(vec![
            Expression::int(0),
            Expression::int(0),
        ])))),
    ])
}

fn binary_varint(buf: Expression) -> Expression {
    binary_uvarint(buf, true)
}

fn decode_signed_varint(ux: Expression) -> Expression {
    Expression::new(ExprKind::Ternary {
        cond: Box::new(binary(
            BinOp::NotEq,
            binary(BinOp::BitAnd, ux.clone(), Expression::int(1)),
            Expression::int(0),
        )),
        then: Box::new(binary(
            BinOp::Mul,
            binary(
                BinOp::Div,
                binary(BinOp::Add, ux.clone(), Expression::int(1)),
                Expression::int(2),
            ),
            Expression::int(-1),
        )),
        else_: Box::new(binary(BinOp::Div, ux, Expression::int(2))),
    })
}

fn put_uint64(little: bool, buf: Expression, value: Expression) -> Expression {
    put_uint64_expr(Expression::bool(little), buf, value)
}

fn put_uint64_expr(order: Expression, buf: Expression, value: Expression) -> Expression {
    let hi = binary(BinOp::Div, value.clone(), Expression::int(4294967296));
    let lo = binary(
        BinOp::Sub,
        value.clone(),
        binary(BinOp::Mul, hi.clone(), Expression::int(4294967296)),
    );
    call(
        "__go_emit_binary_PutUint64PartsWrap",
        vec![order, buf, hi, lo],
    )
}

fn uint64_from_bytes(little: bool, buf: Expression) -> Expression {
    uint64_from_bytes_expr(Expression::bool(little), buf)
}

fn uint64_from_bytes_expr(order: Expression, buf: Expression) -> Expression {
    Expression::new(ExprKind::Ternary {
        cond: Box::new(order),
        then: Box::new(binary(
            BinOp::Add,
            binary(
                BinOp::Add,
                binary(
                    BinOp::Add,
                    index(buf.clone(), Expression::int(0)),
                    binary(
                        BinOp::Shl,
                        index(buf.clone(), Expression::int(1)),
                        Expression::int(8),
                    ),
                ),
                binary(
                    BinOp::Shl,
                    index(buf.clone(), Expression::int(2)),
                    Expression::int(16),
                ),
            ),
            binary(
                BinOp::Add,
                binary(
                    BinOp::Shl,
                    index(buf.clone(), Expression::int(3)),
                    Expression::int(24),
                ),
                binary(
                    BinOp::Mul,
                    uint32_from_be_tail(buf.clone()),
                    Expression::int(4294967296),
                ),
            ),
        )),
        else_: Box::new(binary(
            BinOp::Add,
            binary(
                BinOp::Mul,
                uint32_from_be_head(buf.clone()),
                Expression::int(4294967296),
            ),
            uint32_from_be_tail(buf),
        )),
    })
}

fn uint32_from_be_head(buf: Expression) -> Expression {
    binary(
        BinOp::Add,
        binary(
            BinOp::Add,
            binary(
                BinOp::Shl,
                index(buf.clone(), Expression::int(0)),
                Expression::int(24),
            ),
            binary(
                BinOp::Shl,
                index(buf.clone(), Expression::int(1)),
                Expression::int(16),
            ),
        ),
        binary(
            BinOp::Add,
            binary(
                BinOp::Shl,
                index(buf.clone(), Expression::int(2)),
                Expression::int(8),
            ),
            index(buf, Expression::int(3)),
        ),
    )
}

fn uint32_from_be_tail(buf: Expression) -> Expression {
    binary(
        BinOp::Add,
        binary(
            BinOp::Add,
            binary(
                BinOp::Shl,
                index(buf.clone(), Expression::int(4)),
                Expression::int(24),
            ),
            binary(
                BinOp::Shl,
                index(buf.clone(), Expression::int(5)),
                Expression::int(16),
            ),
        ),
        binary(
            BinOp::Add,
            binary(
                BinOp::Shl,
                index(buf.clone(), Expression::int(6)),
                Expression::int(8),
            ),
            index(buf, Expression::int(7)),
        ),
    )
}

fn emit_order_call(name: &str, little: bool, mut args: Vec<Expression>) -> Expression {
    let mut values = vec![Expression::bool(little)];
    values.append(&mut args);
    call(name, values)
}

fn emit_order_expr_call(name: &str, order: Expression, mut args: Vec<Expression>) -> Expression {
    let mut values = vec![order];
    values.append(&mut args);
    call(name, values)
}

fn copy_bytes_loop(dst_name: &str, src_name: &str) -> Statement {
    Statement::new(StmtKind::For {
        init: Some(Box::new(var_decl("__go_copy_i", Expression::int(0)))),
        cond: Some(binary(
            BinOp::Lt,
            Expression::ident("__go_copy_i"),
            call("len", vec![Expression::ident(src_name)]),
        )),
        update: Some(Expression::new(ExprKind::Assign {
            target: Box::new(Expression::ident("__go_copy_i")),
            value: Box::new(binary(
                BinOp::Add,
                Expression::ident("__go_copy_i"),
                Expression::int(1),
            )),
        })),
        body: vec![assign(
            index(
                Expression::ident(dst_name),
                Expression::ident("__go_copy_i"),
            ),
            index(
                Expression::ident(src_name),
                Expression::ident("__go_copy_i"),
            ),
        )],
    })
}

fn byte_order(little: bool) -> Expression {
    typed_composite(Expression::bool(little), "__goByteOrder")
}

fn base64_encoding(raw: bool, url: bool) -> Expression {
    typed_composite(
        object(vec![
            ("raw", Expression::bool(raw)),
            ("url", Expression::bool(url)),
        ]),
        "__goBase64Encoding",
    )
}

fn typed_composite(expr: Expression, type_name: &str) -> Expression {
    Expression::new(ExprKind::Cast {
        expr: Box::new(expr),
        type_name: type_name.to_string(),
    })
}

fn object(fields: Vec<(&str, Expression)>) -> Expression {
    Expression::new(ExprKind::Object(
        fields
            .into_iter()
            .map(|(key, value)| ObjectProperty::KeyValue {
                key: Expression::string(key),
                value,
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

fn call_method(receiver: Expression, field: &str, args: Vec<Expression>) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(member(receiver, field)),
        args: args.into_iter().map(Argument::positional).collect(),
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

fn arg(args: &[Argument], idx: usize) -> Expression {
    args.get(idx)
        .map(|arg| arg.value.clone())
        .unwrap_or_else(Expression::null)
}

fn arg_values(args: &[Argument]) -> Vec<Expression> {
    args.iter().map(|arg| arg.value.clone()).collect()
}

fn register_type(
    root: &mut Subtree,
    name: &str,
    methods: &[&str],
    member_returns: &[(&str, &str)],
) {
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
            member_returns: member_returns
                .iter()
                .map(|(name, ty)| (name.to_string(), ty.to_string()))
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

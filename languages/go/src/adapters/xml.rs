use vybe_ast::{Argument, BinOp, ExprKind, Expression, PlaceExpr, Statement, StmtKind, UnaryOp};
use vybe_compiler::primitives::namespaces::{NamespaceNode, Subtree};
use vybe_runtime::Value;

pub(crate) fn register_tree(root: &mut Subtree) {
    for (name, emit) in [
        ("encoding.xml.Escape", "go.encoding.xml.Escape"),
        ("encoding.xml.EscapeText", "go.encoding.xml.EscapeText"),
        ("encoding.xml.Unescape", "go.encoding.xml.Unescape"),
        ("encoding.xml.CharData", "go.encoding.xml.CharData"),
        ("encoding.xml.Comment", "go.encoding.xml.Comment"),
        ("encoding.xml.Directive", "go.encoding.xml.Directive"),
        ("encoding.xml.NewDecoder", "go.encoding.xml.NewDecoder"),
        ("encoding.xml.NewEncoder", "go.encoding.xml.NewEncoder"),
        ("encoding.xml.Marshal", "go.encoding.xml.Marshal"),
        (
            "encoding.xml.MarshalIndent",
            "go.encoding.xml.MarshalIndent",
        ),
        ("encoding.xml.Unmarshal", "go.encoding.xml.Unmarshal"),
        ("encoding.xml.Copy", "go.encoding.xml.Copy"),
    ] {
        insert_path(root, name, NamespaceNode::CommonEmit(emit.to_string()));
    }

    insert_path(
        root,
        "encoding.xml.Header",
        NamespaceNode::Const(Value::String(
            "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n".into(),
        )),
    );

    register_type(root, "encoding.xml.Name", &[]);
    register_type(root, "encoding.xml.StartElement", &[]);
    register_type(root, "encoding.xml.EndElement", &[]);
    register_type(root, "encoding.xml.ProcInst", &[]);
    register_type(
        root,
        "encoding.xml.Decoder",
        &[
            "Token",
            "RawToken",
            "Skip",
            "Decode",
            "InputOffset",
            "InputPos",
        ],
    );
    register_type(root, "encoding.xml.Encoder", &["Indent", "Encode"]);
}

pub(crate) fn type_binding(type_name: &str) -> Option<&'static str> {
    match type_name.trim().trim_start_matches('*') {
        "xml.Name" | "encoding/xml.Name" | "encoding.xml.Name" => Some("__goXMLName"),
        "xml.CharData" | "encoding/xml.CharData" | "encoding.xml.CharData" => Some("[]byte"),
        "xml.StartElement" | "encoding/xml.StartElement" | "encoding.xml.StartElement" => {
            Some("__goXMLStartElement")
        }
        "xml.EndElement" | "encoding/xml.EndElement" | "encoding.xml.EndElement" => {
            Some("__goXMLEndElement")
        }
        "xml.ProcInst" | "encoding/xml.ProcInst" | "encoding.xml.ProcInst" => {
            Some("__goXMLProcInst")
        }
        "xml.Decoder" | "encoding/xml.Decoder" | "encoding.xml.Decoder" => Some("__goXMLDecoder"),
        "xml.Encoder" | "encoding/xml.Encoder" | "encoding.xml.Encoder" => Some("__goXMLEncoder"),
        "xml.Comment" | "encoding/xml.Comment" | "encoding.xml.Comment" => Some("[]byte"),
        "xml.Directive" | "encoding/xml.Directive" | "encoding.xml.Directive" => Some("[]byte"),
        _ => None,
    }
}

pub(crate) fn rewrite_member(field: &str) -> Option<Expression> {
    match field {
        "Header" => Some(Expression::string(
            "<?xml version=\"1.0\" encoding=\"UTF-8\"?>\n",
        )),
        _ => None,
    }
}

pub(crate) fn rewrite_simple_call(call_name: &str, args: &[Argument]) -> Option<Expression> {
    match call_name {
        "xml.Escape" | "encoding.xml.Escape" | "go.encoding.xml.Escape" => Some(escape_text(args)),
        "xml.EscapeText" | "encoding.xml.EscapeText" | "go.encoding.xml.EscapeText" => {
            Some(escape_text(args))
        }
        "xml.Unescape" | "encoding.xml.Unescape" | "go.encoding.xml.Unescape" => {
            Some(Expression::new(ExprKind::Tuple(vec![
                unescape_string(bytes_to_string(arg(args, 0))),
                Expression::null(),
            ])))
        }
        "xml.CharData" | "encoding.xml.CharData" | "go.encoding.xml.CharData" => Some(arg(args, 0)),
        "xml.Comment"
        | "encoding.xml.Comment"
        | "go.encoding.xml.Comment"
        | "xml.Directive"
        | "encoding.xml.Directive"
        | "go.encoding.xml.Directive" => Some(arg(args, 0)),
        "xml.NewDecoder" | "encoding.xml.NewDecoder" | "go.encoding.xml.NewDecoder" => Some(
            pointer_arg(decoder_object(reader_text(value_arg(arg(args, 0))))),
        ),
        "xml.NewEncoder" | "encoding.xml.NewEncoder" | "go.encoding.xml.NewEncoder" => {
            Some(pointer_arg(encoder_object(writer_value(arg(args, 0)))))
        }
        "xml.Copy" | "encoding.xml.Copy" | "go.encoding.xml.Copy" => Some(copy_expr(args)),
        _ => None,
    }
}

pub(crate) fn rewrite_method_call(
    object: Expression,
    receiver_type: &str,
    field: &str,
    args: &[Argument],
) -> Option<Expression> {
    let receiver = if receiver_type.trim().starts_with('*') {
        Expression::new(ExprKind::RefLoad(Box::new(object)))
    } else {
        object
    };
    match (receiver_type.trim().trim_start_matches('*').trim(), field) {
        ("__goXMLDecoder", "Token" | "RawToken") => Some(decoder_token(receiver)),
        ("__goXMLDecoder", "Skip" | "Decode") => Some(Expression::null()),
        ("__goXMLDecoder", "InputOffset") => Some(member(receiver, "pos")),
        ("__goXMLDecoder", "InputPos") => Some(Expression::new(ExprKind::Tuple(vec![
            Expression::int(1),
            binary(BinOp::Add, member(receiver, "pos"), Expression::int(1)),
        ]))),
        ("__goXMLEncoder", "Indent") => Some(encoder_indent(receiver, args)),
        ("__goXMLEncoder", "Flush" | "Close") => Some(Expression::null()),
        _ => None,
    }
}

pub(crate) fn encoder_encode(receiver: Expression, text: Expression) -> Expression {
    let receiver = value_arg(receiver);
    lambda_call(vec![
        expr_stmt(crate::adapters::bytes_io::buffer_write_string_expr(
            member(receiver, "w"),
            text,
        )),
        Statement::new(StmtKind::Return(Some(Expression::null()))),
    ])
}

pub(crate) fn name_expr(
    namespace: Expression,
    local: Expression,
    prefix: Expression,
) -> Expression {
    call("__go_xml_name", vec![namespace, local, prefix])
}

pub(crate) fn name_local_expr(expr: Expression) -> Expression {
    call("__go_xml_name_local", vec![expr])
}

pub(crate) fn name_space_expr(expr: Expression) -> Expression {
    call("__go_xml_name_space", vec![expr])
}

pub(crate) fn token_kind_expr(expr: Expression) -> Expression {
    member(expr, "Kind")
}

pub(crate) fn token_local_expr(expr: Expression) -> Expression {
    name_local_expr(member(expr, "Name"))
}

pub(crate) fn escape_string(expr: Expression) -> Expression {
    let replacements = [
        ("&", "&amp;"),
        ("<", "&lt;"),
        (">", "&gt;"),
        ("\"", "&quot;"),
        ("'", "&apos;"),
    ];
    replacements
        .into_iter()
        .fold(call("__go_fmt_string", vec![expr]), |acc, (old, new)| {
            call(
                "strings.ReplaceAll",
                vec![acc, Expression::string(old), Expression::string(new)],
            )
        })
}

pub(crate) fn unescape_string(expr: Expression) -> Expression {
    let replacements = [
        ("&lt;", "<"),
        ("&gt;", ">"),
        ("&quot;", "\""),
        ("&apos;", "'"),
        ("&amp;", "&"),
    ];
    replacements
        .into_iter()
        .fold(call("__go_fmt_string", vec![expr]), |acc, (old, new)| {
            call(
                "strings.ReplaceAll",
                vec![acc, Expression::string(old), Expression::string(new)],
            )
        })
}

pub(crate) fn attr_expr(src: Expression, name: Expression) -> Expression {
    lambda_call(vec![
        var_decl("__go_xml_s", any_string(src)),
        var_decl(
            "__go_xml_needle",
            binary(BinOp::Add, name, Expression::string("=\"")),
        ),
        var_decl(
            "__go_xml_i",
            call(
                "strings.Index",
                vec![
                    Expression::ident("__go_xml_s"),
                    Expression::ident("__go_xml_needle"),
                ],
            ),
        ),
        Statement::new(StmtKind::If {
            cond: binary(
                BinOp::Lt,
                Expression::ident("__go_xml_i"),
                Expression::int(0),
            ),
            then_body: vec![Statement::new(StmtKind::Return(Some(Expression::string(
                "",
            ))))],
            elifs: Vec::new(),
            else_body: None,
        }),
        var_decl(
            "__go_xml_start",
            binary(
                BinOp::Add,
                Expression::ident("__go_xml_i"),
                call("len", vec![Expression::ident("__go_xml_needle")]),
            ),
        ),
        var_decl(
            "__go_xml_tail",
            string_slice(
                Expression::ident("__go_xml_s"),
                Expression::ident("__go_xml_start"),
                call("len", vec![Expression::ident("__go_xml_s")]),
            ),
        ),
        var_decl(
            "__go_xml_end",
            call(
                "strings.Index",
                vec![Expression::ident("__go_xml_tail"), Expression::string("\"")],
            ),
        ),
        Statement::new(StmtKind::If {
            cond: binary(
                BinOp::Lt,
                Expression::ident("__go_xml_end"),
                Expression::int(0),
            ),
            then_body: vec![Statement::new(StmtKind::Return(Some(Expression::string(
                "",
            ))))],
            elifs: Vec::new(),
            else_body: None,
        }),
        Statement::new(StmtKind::Return(Some(unescape_string(string_slice(
            Expression::ident("__go_xml_tail"),
            Expression::int(0),
            Expression::ident("__go_xml_end"),
        ))))),
    ])
}

pub(crate) fn elem_expr(src: Expression, name: Expression) -> Expression {
    lambda_call(vec![
        var_decl("__go_xml_s", any_string(src)),
        var_decl(
            "__go_xml_open",
            binary(BinOp::Add, Expression::string("<"), name.clone()),
        ),
        var_decl(
            "__go_xml_i",
            call(
                "strings.Index",
                vec![
                    Expression::ident("__go_xml_s"),
                    Expression::ident("__go_xml_open"),
                ],
            ),
        ),
        Statement::new(StmtKind::If {
            cond: binary(
                BinOp::Lt,
                Expression::ident("__go_xml_i"),
                Expression::int(0),
            ),
            then_body: vec![Statement::new(StmtKind::Return(Some(Expression::string(
                "",
            ))))],
            elifs: Vec::new(),
            else_body: None,
        }),
        var_decl(
            "__go_xml_after",
            string_slice(
                Expression::ident("__go_xml_s"),
                binary(
                    BinOp::Add,
                    Expression::ident("__go_xml_i"),
                    call("len", vec![Expression::ident("__go_xml_open")]),
                ),
                call("len", vec![Expression::ident("__go_xml_s")]),
            ),
        ),
        var_decl(
            "__go_xml_gt",
            call(
                "strings.Index",
                vec![Expression::ident("__go_xml_after"), Expression::string(">")],
            ),
        ),
        Statement::new(StmtKind::If {
            cond: binary(
                BinOp::Lt,
                Expression::ident("__go_xml_gt"),
                Expression::int(0),
            ),
            then_body: vec![Statement::new(StmtKind::Return(Some(Expression::string(
                "",
            ))))],
            elifs: Vec::new(),
            else_body: None,
        }),
        var_decl(
            "__go_xml_body",
            string_slice(
                Expression::ident("__go_xml_after"),
                binary(
                    BinOp::Add,
                    Expression::ident("__go_xml_gt"),
                    Expression::int(1),
                ),
                call("len", vec![Expression::ident("__go_xml_after")]),
            ),
        ),
        var_decl(
            "__go_xml_close",
            binary(
                BinOp::Add,
                binary(BinOp::Add, Expression::string("</"), name),
                Expression::string(">"),
            ),
        ),
        var_decl(
            "__go_xml_end",
            call(
                "strings.Index",
                vec![
                    Expression::ident("__go_xml_body"),
                    Expression::ident("__go_xml_close"),
                ],
            ),
        ),
        Statement::new(StmtKind::If {
            cond: binary(
                BinOp::Lt,
                Expression::ident("__go_xml_end"),
                Expression::int(0),
            ),
            then_body: vec![Statement::new(StmtKind::Return(Some(Expression::string(
                "",
            ))))],
            elifs: Vec::new(),
            else_body: None,
        }),
        Statement::new(StmtKind::Return(Some(unescape_string(string_slice(
            Expression::ident("__go_xml_body"),
            Expression::int(0),
            Expression::ident("__go_xml_end"),
        ))))),
    ])
}

pub(crate) fn chardata_expr(src: Expression) -> Expression {
    lambda_call(vec![
        var_decl("__go_xml_s", any_string(src)),
        var_decl(
            "__go_xml_start",
            call(
                "strings.Index",
                vec![Expression::ident("__go_xml_s"), Expression::string(">")],
            ),
        ),
        Statement::new(StmtKind::If {
            cond: binary(
                BinOp::Lt,
                Expression::ident("__go_xml_start"),
                Expression::int(0),
            ),
            then_body: vec![Statement::new(StmtKind::Return(Some(Expression::string(
                "",
            ))))],
            elifs: Vec::new(),
            else_body: None,
        }),
        var_decl(
            "__go_xml_tail",
            string_slice(
                Expression::ident("__go_xml_s"),
                binary(
                    BinOp::Add,
                    Expression::ident("__go_xml_start"),
                    Expression::int(1),
                ),
                call("len", vec![Expression::ident("__go_xml_s")]),
            ),
        ),
        var_decl(
            "__go_xml_end",
            call(
                "strings.Index",
                vec![Expression::ident("__go_xml_tail"), Expression::string("<")],
            ),
        ),
        Statement::new(StmtKind::If {
            cond: binary(
                BinOp::Lt,
                Expression::ident("__go_xml_end"),
                Expression::int(0),
            ),
            then_body: vec![Statement::new(StmtKind::Return(Some(Expression::string(
                "",
            ))))],
            elifs: Vec::new(),
            else_body: None,
        }),
        Statement::new(StmtKind::Return(Some(unescape_string(string_slice(
            Expression::ident("__go_xml_tail"),
            Expression::int(0),
            Expression::ident("__go_xml_end"),
        ))))),
    ])
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

fn escape_text(args: &[Argument]) -> Expression {
    lambda_call(vec![
        expr_stmt(crate::adapters::bytes_io::buffer_write_string_expr(
            writer_value(arg(args, 0)),
            escape_string(bytes_to_string(arg(args, 1))),
        )),
        Statement::new(StmtKind::Return(Some(Expression::null()))),
    ])
}

fn copy_expr(args: &[Argument]) -> Expression {
    lambda_call(vec![
        expr_stmt(crate::adapters::bytes_io::buffer_write_string_expr(
            writer_value(arg(args, 0)),
            readable_string(arg(args, 1)),
        )),
        Statement::new(StmtKind::Return(Some(Expression::null()))),
    ])
}

fn decoder_object(data: Expression) -> Expression {
    typed_object(
        "__goXMLDecoder",
        vec![
            ("data", data),
            ("pos", Expression::int(0)),
            ("pendingEnd", Expression::string("")),
            ("Entity", Expression::new(ExprKind::Object(Vec::new()))),
        ],
    )
}

fn encoder_object(writer: Expression) -> Expression {
    typed_object(
        "__goXMLEncoder",
        vec![
            ("w", writer),
            ("prefix", Expression::string("")),
            ("indent", Expression::string("")),
        ],
    )
}

fn decoder_token(receiver: Expression) -> Expression {
    let d = Expression::ident("__go_xml_d");
    lambda_call(vec![
        var_decl("__go_xml_d", receiver),
        Statement::new(StmtKind::If {
            cond: binary(
                BinOp::NotEq,
                member(d.clone(), "pendingEnd"),
                Expression::string(""),
            ),
            then_body: vec![
                var_decl("__go_xml_tag", member(d.clone(), "pendingEnd")),
                expr_stmt(Expression::new(ExprKind::Assign {
                    target: Box::new(member(d.clone(), "pendingEnd")),
                    value: Box::new(Expression::string("")),
                })),
                Statement::new(StmtKind::Return(Some(Expression::new(ExprKind::Tuple(
                    vec![
                        end_element(Expression::ident("__go_xml_tag")),
                        Expression::null(),
                    ],
                ))))),
            ],
            elifs: Vec::new(),
            else_body: None,
        }),
        Statement::new(StmtKind::If {
            cond: binary(
                BinOp::GtEq,
                member(d.clone(), "pos"),
                call("len", vec![member(d.clone(), "data")]),
            ),
            then_body: vec![Statement::new(StmtKind::Return(Some(Expression::new(
                ExprKind::Tuple(vec![Expression::null(), Expression::string("EOF")]),
            ))))],
            elifs: Vec::new(),
            else_body: None,
        }),
        Statement::new(StmtKind::If {
            cond: binary(
                BinOp::NotEq,
                string_slice(
                    member(d.clone(), "data"),
                    member(d.clone(), "pos"),
                    binary(BinOp::Add, member(d.clone(), "pos"), Expression::int(1)),
                ),
                Expression::string("<"),
            ),
            then_body: vec![
                var_decl("__go_xml_start", member(d.clone(), "pos")),
                var_decl(
                    "__go_xml_next",
                    call(
                        "strings.Index",
                        vec![
                            string_slice(
                                member(d.clone(), "data"),
                                member(d.clone(), "pos"),
                                call("len", vec![member(d.clone(), "data")]),
                            ),
                            Expression::string("<"),
                        ],
                    ),
                ),
                Statement::new(StmtKind::If {
                    cond: binary(
                        BinOp::Lt,
                        Expression::ident("__go_xml_next"),
                        Expression::int(0),
                    ),
                    then_body: vec![expr_stmt(Expression::new(ExprKind::Assign {
                        target: Box::new(member(d.clone(), "pos")),
                        value: Box::new(call("len", vec![member(d.clone(), "data")])),
                    }))],
                    elifs: Vec::new(),
                    else_body: Some(vec![expr_stmt(Expression::new(ExprKind::Assign {
                        target: Box::new(member(d.clone(), "pos")),
                        value: Box::new(binary(
                            BinOp::Add,
                            Expression::ident("__go_xml_start"),
                            Expression::ident("__go_xml_next"),
                        )),
                    }))]),
                }),
                Statement::new(StmtKind::Return(Some(Expression::new(ExprKind::Tuple(
                    vec![
                        string_to_bytes(string_slice(
                            member(d.clone(), "data"),
                            Expression::ident("__go_xml_start"),
                            member(d.clone(), "pos"),
                        )),
                        Expression::null(),
                    ],
                ))))),
            ],
            elifs: Vec::new(),
            else_body: None,
        }),
        var_decl(
            "__go_xml_close_rel",
            call(
                "strings.Index",
                vec![
                    string_slice(
                        member(d.clone(), "data"),
                        member(d.clone(), "pos"),
                        call("len", vec![member(d.clone(), "data")]),
                    ),
                    Expression::string(">"),
                ],
            ),
        ),
        Statement::new(StmtKind::If {
            cond: binary(
                BinOp::Lt,
                Expression::ident("__go_xml_close_rel"),
                Expression::int(0),
            ),
            then_body: vec![
                expr_stmt(Expression::new(ExprKind::Assign {
                    target: Box::new(member(d.clone(), "pos")),
                    value: Box::new(call("len", vec![member(d.clone(), "data")])),
                })),
                Statement::new(StmtKind::Return(Some(Expression::new(ExprKind::Tuple(
                    vec![Expression::null(), Expression::string("EOF")],
                ))))),
            ],
            elifs: Vec::new(),
            else_body: None,
        }),
        var_decl(
            "__go_xml_close",
            binary(
                BinOp::Add,
                member(d.clone(), "pos"),
                Expression::ident("__go_xml_close_rel"),
            ),
        ),
        var_decl(
            "__go_xml_tag",
            string_slice(
                member(d.clone(), "data"),
                binary(BinOp::Add, member(d.clone(), "pos"), Expression::int(1)),
                Expression::ident("__go_xml_close"),
            ),
        ),
        expr_stmt(Expression::new(ExprKind::Assign {
            target: Box::new(member(d.clone(), "pos")),
            value: Box::new(binary(
                BinOp::Add,
                Expression::ident("__go_xml_close"),
                Expression::int(1),
            )),
        })),
        Statement::new(StmtKind::If {
            cond: binary(
                BinOp::Eq,
                string_slice(
                    Expression::ident("__go_xml_tag"),
                    Expression::int(0),
                    Expression::int(1),
                ),
                Expression::string("/"),
            ),
            then_body: vec![Statement::new(StmtKind::Return(Some(Expression::new(
                ExprKind::Tuple(vec![
                    end_element(string_slice(
                        Expression::ident("__go_xml_tag"),
                        Expression::int(1),
                        call("len", vec![Expression::ident("__go_xml_tag")]),
                    )),
                    Expression::null(),
                ]),
            ))))],
            elifs: Vec::new(),
            else_body: None,
        }),
        var_decl(
            "__go_xml_self_closing",
            binary(
                BinOp::GtEq,
                call(
                    "strings.Index",
                    vec![Expression::ident("__go_xml_tag"), Expression::string("/")],
                ),
                Expression::int(0),
            ),
        ),
        expr_stmt(Expression::new(ExprKind::Assign {
            target: Box::new(Expression::ident("__go_xml_tag")),
            value: Box::new(call(
                "strings.ReplaceAll",
                vec![
                    Expression::ident("__go_xml_tag"),
                    Expression::string("/"),
                    Expression::string(""),
                ],
            )),
        })),
        var_decl(
            "__go_xml_space",
            call(
                "strings.Index",
                vec![Expression::ident("__go_xml_tag"), Expression::string(" ")],
            ),
        ),
        Statement::new(StmtKind::If {
            cond: binary(
                BinOp::GtEq,
                Expression::ident("__go_xml_space"),
                Expression::int(0),
            ),
            then_body: vec![expr_stmt(Expression::new(ExprKind::Assign {
                target: Box::new(Expression::ident("__go_xml_tag")),
                value: Box::new(string_slice(
                    Expression::ident("__go_xml_tag"),
                    Expression::int(0),
                    Expression::ident("__go_xml_space"),
                )),
            }))],
            elifs: Vec::new(),
            else_body: None,
        }),
        Statement::new(StmtKind::If {
            cond: Expression::ident("__go_xml_self_closing"),
            then_body: vec![expr_stmt(Expression::new(ExprKind::Assign {
                target: Box::new(member(d, "pendingEnd")),
                value: Box::new(Expression::ident("__go_xml_tag")),
            }))],
            elifs: Vec::new(),
            else_body: None,
        }),
        Statement::new(StmtKind::Return(Some(Expression::new(ExprKind::Tuple(
            vec![
                start_element(Expression::ident("__go_xml_tag")),
                Expression::null(),
            ],
        ))))),
    ])
}

fn start_element(tag: Expression) -> Expression {
    typed_object(
        "__goXMLStartElement",
        vec![
            (
                "Name",
                name_expr(Expression::string(""), tag.clone(), Expression::string("")),
            ),
            ("Kind", Expression::string("start")),
            ("Tag", tag),
        ],
    )
}

fn end_element(tag: Expression) -> Expression {
    typed_object(
        "__goXMLEndElement",
        vec![
            (
                "Name",
                name_expr(Expression::string(""), tag.clone(), Expression::string("")),
            ),
            ("Kind", Expression::string("end")),
            ("Tag", tag),
        ],
    )
}

fn encoder_indent(receiver: Expression, args: &[Argument]) -> Expression {
    lambda_call(vec![
        expr_stmt(Expression::new(ExprKind::Assign {
            target: Box::new(member(receiver.clone(), "prefix")),
            value: Box::new(arg(args, 0)),
        })),
        expr_stmt(Expression::new(ExprKind::Assign {
            target: Box::new(member(receiver, "indent")),
            value: Box::new(arg(args, 1)),
        })),
        Statement::new(StmtKind::Return(Some(Expression::null()))),
    ])
}

fn writer_value(expr: Expression) -> Expression {
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

fn readable_string(expr: Expression) -> Expression {
    let value = value_arg(expr);
    Expression::new(ExprKind::Ternary {
        cond: Box::new(binary(
            BinOp::NotEq,
            member(value.clone(), "data"),
            Expression::null(),
        )),
        then: Box::new(crate::adapters::bytes_io::buffer_string(value.clone())),
        else_: Box::new(reader_text(value)),
    })
}

fn any_string(expr: Expression) -> Expression {
    Expression::new(ExprKind::Ternary {
        cond: Box::new(binary(
            BinOp::NotEq,
            member(expr.clone(), "data"),
            Expression::null(),
        )),
        then: Box::new(expr.clone()),
        else_: Box::new(bytes_to_string(expr)),
    })
}

fn bytes_to_string(expr: Expression) -> Expression {
    crate::adapters::bytes_io::bytes_to_string(expr)
}

fn string_to_bytes(expr: Expression) -> Expression {
    crate::adapters::bytes_io::string_to_bytes(expr)
}

fn reader_text(expr: Expression) -> Expression {
    crate::adapters::bytes_io::reader_text(expr)
}

fn typed_object(type_name: &str, fields: Vec<(&str, Expression)>) -> Expression {
    crate::adapters::bytes_io::typed_object(type_name, fields)
}

fn pointer_arg(expr: Expression) -> Expression {
    crate::adapters::bytes_io::pointer_arg(expr)
}

fn value_arg(expr: Expression) -> Expression {
    crate::adapters::bytes_io::value_arg(expr)
}

fn member(object: Expression, field: &str) -> Expression {
    crate::adapters::bytes_io::member(object, field)
}

fn string_slice(object: Expression, start: Expression, end: Expression) -> Expression {
    crate::adapters::bytes_io::string_slice(object, start, end)
}

fn binary(op: BinOp, left: Expression, right: Expression) -> Expression {
    crate::adapters::bytes_io::binary(op, left, right)
}

fn var_decl(name: &str, init: Expression) -> Statement {
    crate::adapters::bytes_io::var_decl(name, init)
}

fn expr_stmt(expr: Expression) -> Statement {
    crate::adapters::bytes_io::expr_stmt(expr)
}

fn lambda_call(body: Vec<Statement>) -> Expression {
    crate::adapters::bytes_io::lambda_call(body)
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

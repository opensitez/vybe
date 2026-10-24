//! C string.h runtime helpers (libc surface): the `__c_str*_h` functions backing
//! strcoll / strxfrm / strpbrk / strspn / strcspn. The call-site lowerings live
//! in the walker (they map the C name to these helpers); the helper bodies live
//! here so any libc-targeting front-end can inject them. Injected once into the
//! program prelude.

use crate::emitter::build::*;
use vybe_ast::{BinOp, ExprKind, Statement, StmtKind};

pub fn runtime_helpers() -> Vec<Statement> {
    let mut out: Vec<Statement> = Vec::new();
    out.push(c_source_byte());
    out.push(c_memccpy());
    out.push(c_utf8_store());
    out.push(c_wide_store());
    out.push(var_decl_stmt("__c_strtok_save", null_lit()));
    out.push(c_strtok_is_delimiter());
    out.push(c_strtok());
    out.push(c_strstr_linear());
    out.push(function_stmt(
        "__libc_memset",
        vec!["dst", "ch", "n"],
        vec![
            var_decl_stmt("i", int_lit(0)),
            stmt(StmtKind::While {
                cond: expr(ExprKind::Binary {
                    op: BinOp::Lt,
                    left: Box::new(ident("i")),
                    right: Box::new(ident("n")),
                }),
                body: vec![
                    stmt(StmtKind::Expr(assign_expr(
                        index_expr(ident("dst"), ident("i")),
                        ident("ch"),
                    ))),
                    stmt(StmtKind::Expr(assign_expr(
                        ident("i"),
                        expr(ExprKind::Binary {
                            op: BinOp::Add,
                            left: Box::new(ident("i")),
                            right: Box::new(int_lit(1)),
                        }),
                    ))),
                ],
                else_body: None,
            }),
            stmt(StmtKind::Return(Some(ident("dst")))),
        ],
    ));
    out.push(typed_function_stmt(
        "__c_mem_index_of_h",
        vec!["buf", "needle", "n", "reverse"],
        "int64",
        vec![
            // The caller supplies bytes, not a managed-container length.
            var_decl_stmt("limit", ident("n")),
            var_decl_stmt(
                "needle_byte",
                expr(ExprKind::Binary {
                    op: BinOp::Integer(vybe_ast::IntOp::And, vybe_ast::BitLane::W32),
                    left: Box::new(expr(ExprKind::Ternary {
                        cond: Box::new(expr(ExprKind::IsType {
                            expr: Box::new(ident("needle")),
                            type_name: "string".into(),
                        })),
                        then: Box::new(call_expr(
                            ident("__c_char_code_at"),
                            vec![ident("needle"), int_lit(0)],
                        )),
                        else_: Box::new(ident("needle")),
                    })),
                    right: Box::new(int_lit(255)),
                }),
            ),
            var_decl_stmt(
                "i",
                expr(ExprKind::Ternary {
                    cond: Box::new(ident("reverse")),
                    then: Box::new(expr(ExprKind::Binary {
                        op: BinOp::Sub,
                        left: Box::new(ident("limit")),
                        right: Box::new(int_lit(1)),
                    })),
                    else_: Box::new(int_lit(0)),
                }),
            ),
            stmt(StmtKind::While {
                cond: expr(ExprKind::Ternary {
                    cond: Box::new(ident("reverse")),
                    then: Box::new(expr(ExprKind::Binary {
                        op: BinOp::GtEq,
                        left: Box::new(ident("i")),
                        right: Box::new(int_lit(0)),
                    })),
                    else_: Box::new(expr(ExprKind::Binary {
                        op: BinOp::Lt,
                        left: Box::new(ident("i")),
                        right: Box::new(ident("limit")),
                    })),
                }),
                body: vec![
                    var_decl_stmt(
                        "value",
                        call_expr(ident("__c_char_ptr_read"), vec![ident("buf"), ident("i")]),
                    ),
                    if_stmt(
                        expr(ExprKind::Binary {
                            op: BinOp::Integer(vybe_ast::IntOp::Eq, vybe_ast::BitLane::W32),
                            left: Box::new(expr(ExprKind::Binary {
                                op: BinOp::Integer(vybe_ast::IntOp::And, vybe_ast::BitLane::W32),
                                left: Box::new(ident("value")),
                                right: Box::new(int_lit(255)),
                            })),
                            right: Box::new(ident("needle_byte")),
                        }),
                        vec![stmt(StmtKind::Return(Some(ident("i"))))],
                        None,
                    ),
                    stmt(StmtKind::Expr(assign_expr(
                        ident("i"),
                        expr(ExprKind::Ternary {
                            cond: Box::new(ident("reverse")),
                            then: Box::new(expr(ExprKind::Binary {
                                op: BinOp::Sub,
                                left: Box::new(ident("i")),
                                right: Box::new(int_lit(1)),
                            })),
                            else_: Box::new(expr(ExprKind::Binary {
                                op: BinOp::Add,
                                left: Box::new(ident("i")),
                                right: Box::new(int_lit(1)),
                            })),
                        }),
                    ))),
                ],
                else_body: None,
            }),
            stmt(StmtKind::Return(Some(int_lit(-1)))),
        ],
    ));

    out.push(function_stmt(
        "__c_strnlen_h",
        vec!["s", "maxlen"],
        vec![
            var_decl_stmt(
                "limit",
                expr(ExprKind::Ternary {
                    cond: Box::new(expr(ExprKind::Binary {
                        op: BinOp::Lt,
                        left: Box::new(ident("maxlen")),
                        right: Box::new(member(ident("s"), "length")),
                    })),
                    then: Box::new(ident("maxlen")),
                    else_: Box::new(member(ident("s"), "length")),
                }),
            ),
            var_decl_stmt("i", int_lit(0)),
            stmt(StmtKind::While {
                cond: expr(ExprKind::Binary {
                    op: BinOp::Lt,
                    left: Box::new(ident("i")),
                    right: Box::new(ident("limit")),
                }),
                body: vec![
                    var_decl_stmt("value", index_expr(ident("s"), ident("i"))),
                    if_stmt(
                        expr(ExprKind::Binary {
                            op: BinOp::Or,
                            left: Box::new(expr(ExprKind::Binary {
                                op: BinOp::Eq,
                                left: Box::new(ident("value")),
                                right: Box::new(int_lit(0)),
                            })),
                            right: Box::new(expr(ExprKind::Binary {
                                op: BinOp::Eq,
                                left: Box::new(ident("value")),
                                right: Box::new(str_lit("\0")),
                            })),
                        }),
                        vec![stmt(StmtKind::Return(Some(ident("i"))))],
                        None,
                    ),
                    stmt(StmtKind::Expr(assign_expr(
                        ident("i"),
                        expr(ExprKind::Binary {
                            op: BinOp::Add,
                            left: Box::new(ident("i")),
                            right: Box::new(int_lit(1)),
                        }),
                    ))),
                ],
                else_body: None,
            }),
            stmt(StmtKind::Return(Some(ident("i")))),
        ],
    ));

    out.push(function_stmt(
        "__c_strcoll_h",
        vec!["a", "b"],
        vec![stmt(StmtKind::Return(Some(call_expr(
            ident("strcmp"),
            vec![ident("a"), ident("b")],
        ))))],
    ));

    out.push(function_stmt(
        "__c_strncasecmp8_addr_h",
        vec!["addr", "b"],
        vec![
            var_decl_stmt("i", int_lit(0)),
            var_decl_stmt("ca", int_lit(0)),
            var_decl_stmt("cb", int_lit(0)),
            stmt(StmtKind::While {
                cond: expr(ExprKind::Binary {
                    op: BinOp::Lt,
                    left: Box::new(ident("i")),
                    right: Box::new(int_lit(8)),
                }),
                body: vec![
                    stmt(StmtKind::Expr(assign_expr(
                        ident("ca"),
                        call_expr(
                            ident("__c_ptr_i32_load8_u"),
                            vec![expr(ExprKind::Binary {
                                op: BinOp::Add,
                                left: Box::new(ident("addr")),
                                right: Box::new(ident("i")),
                            })],
                        ),
                    ))),
                    stmt(StmtKind::Expr(assign_expr(
                        ident("cb"),
                        call_expr(ident("__c_char_ptr_read"), vec![ident("b"), ident("i")]),
                    ))),
                    if_stmt(
                        expr(ExprKind::Binary {
                            op: BinOp::And,
                            left: Box::new(expr(ExprKind::Binary {
                                op: BinOp::GtEq,
                                left: Box::new(ident("ca")),
                                right: Box::new(int_lit(65)),
                            })),
                            right: Box::new(expr(ExprKind::Binary {
                                op: BinOp::LtEq,
                                left: Box::new(ident("ca")),
                                right: Box::new(int_lit(90)),
                            })),
                        }),
                        vec![stmt(StmtKind::Expr(assign_expr(
                            ident("ca"),
                            expr(ExprKind::Binary {
                                op: BinOp::Add,
                                left: Box::new(ident("ca")),
                                right: Box::new(int_lit(32)),
                            }),
                        )))],
                        None,
                    ),
                    if_stmt(
                        expr(ExprKind::Binary {
                            op: BinOp::And,
                            left: Box::new(expr(ExprKind::Binary {
                                op: BinOp::GtEq,
                                left: Box::new(ident("cb")),
                                right: Box::new(int_lit(65)),
                            })),
                            right: Box::new(expr(ExprKind::Binary {
                                op: BinOp::LtEq,
                                left: Box::new(ident("cb")),
                                right: Box::new(int_lit(90)),
                            })),
                        }),
                        vec![stmt(StmtKind::Expr(assign_expr(
                            ident("cb"),
                            expr(ExprKind::Binary {
                                op: BinOp::Add,
                                left: Box::new(ident("cb")),
                                right: Box::new(int_lit(32)),
                            }),
                        )))],
                        None,
                    ),
                    if_stmt(
                        expr(ExprKind::Binary {
                            op: BinOp::Or,
                            left: Box::new(expr(ExprKind::Binary {
                                op: BinOp::NotEq,
                                left: Box::new(ident("ca")),
                                right: Box::new(ident("cb")),
                            })),
                            right: Box::new(expr(ExprKind::Binary {
                                op: BinOp::Eq,
                                left: Box::new(ident("ca")),
                                right: Box::new(int_lit(0)),
                            })),
                        }),
                        vec![stmt(StmtKind::Return(Some(expr(ExprKind::Binary {
                            op: BinOp::Sub,
                            left: Box::new(ident("ca")),
                            right: Box::new(ident("cb")),
                        }))))],
                        None,
                    ),
                    stmt(StmtKind::Expr(assign_expr(
                        ident("i"),
                        expr(ExprKind::Binary {
                            op: BinOp::Add,
                            left: Box::new(ident("i")),
                            right: Box::new(int_lit(1)),
                        }),
                    ))),
                ],
                else_body: None,
            }),
            stmt(StmtKind::Return(Some(int_lit(0)))),
        ],
    ));

    out.push(function_stmt(
        "__c_strncasecmp8_ptrvec_field_h",
        vec!["vec", "idx", "field", "offset", "b"],
        vec![
            var_decl_stmt(
                "entry",
                expr(ExprKind::Ternary {
                    cond: Box::new(expr(ExprKind::Binary {
                        op: BinOp::Eq,
                        left: Box::new(expr(ExprKind::Unary {
                            op: vybe_ast::UnaryOp::Typeof,
                            expr: Box::new(ident("vec")),
                        })),
                        right: Box::new(str_lit("number")),
                    })),
                    then: Box::new(call_expr(
                        ident("__c_ptr_i32_load"),
                        vec![expr(ExprKind::Binary {
                            op: BinOp::Add,
                            left: Box::new(ident("vec")),
                            right: Box::new(expr(ExprKind::Binary {
                                op: BinOp::Mul,
                                left: Box::new(ident("idx")),
                                right: Box::new(int_lit(8)),
                            })),
                        })],
                    )),
                    else_: Box::new(index_expr(ident("vec"), ident("idx"))),
                }),
            ),
            var_decl_stmt(
                "addr",
                call_expr(
                    ident("__c_hybrid_struct_field_ptr"),
                    vec![ident("entry"), ident("field"), ident("offset")],
                ),
            ),
            stmt(StmtKind::Return(Some(call_expr(
                ident("__c_strncasecmp8_addr_h"),
                vec![ident("addr"), ident("b")],
            )))),
        ],
    ));

    out.push(function_stmt(
        "__c_strxfrm_h",
        vec!["dst", "src", "n"],
        vec![
            var_decl_stmt(
                "max_len",
                expr(ExprKind::Ternary {
                    cond: Box::new(expr(ExprKind::Binary {
                        op: BinOp::Lt,
                        left: Box::new(ident("n")),
                        right: Box::new(int_lit(1)),
                    })),
                    then: Box::new(int_lit(0)),
                    else_: Box::new(expr(ExprKind::Binary {
                        op: BinOp::Sub,
                        left: Box::new(ident("n")),
                        right: Box::new(int_lit(1)),
                    })),
                }),
            ),
            var_decl_stmt(
                "out",
                call_member(
                    ident("src"),
                    "substring",
                    vec![int_lit(0), ident("max_len")],
                ),
            ),
            stmt(StmtKind::Expr(assign_expr(ident("dst"), ident("out")))),
            stmt(StmtKind::Return(Some(member(ident("src"), "length")))),
        ],
    ));

    out.push(function_stmt(
        "__c_strpbrk_h",
        vec!["s", "accept"],
        vec![
            var_decl_stmt("i", int_lit(0)),
            stmt(StmtKind::While {
                cond: expr(ExprKind::Binary {
                    op: BinOp::And,
                    left: Box::new(expr(ExprKind::Binary {
                        op: BinOp::Lt,
                        left: Box::new(ident("i")),
                        right: Box::new(member(ident("s"), "length")),
                    })),
                    right: Box::new(expr(ExprKind::Binary {
                        op: BinOp::Lt,
                        left: Box::new(call_member(
                            ident("accept"),
                            "indexOf",
                            vec![call_member(ident("s"), "charAt", vec![ident("i")])],
                        )),
                        right: Box::new(int_lit(0)),
                    })),
                }),
                body: vec![stmt(StmtKind::Expr(assign_expr(
                    ident("i"),
                    expr(ExprKind::Binary {
                        op: BinOp::Add,
                        left: Box::new(ident("i")),
                        right: Box::new(int_lit(1)),
                    }),
                )))],
                else_body: None,
            }),
            if_stmt(
                expr(ExprKind::Binary {
                    op: BinOp::GtEq,
                    left: Box::new(ident("i")),
                    right: Box::new(member(ident("s"), "length")),
                }),
                vec![stmt(StmtKind::Return(Some(null_lit())))],
                None,
            ),
            stmt(StmtKind::Return(Some(call_member(
                ident("s"),
                "substring",
                vec![ident("i")],
            )))),
        ],
    ));

    out.push(function_stmt(
        "__c_strspn_h",
        vec!["s", "accept"],
        vec![
            var_decl_stmt("i", int_lit(0)),
            stmt(StmtKind::While {
                cond: expr(ExprKind::Binary {
                    op: BinOp::And,
                    left: Box::new(expr(ExprKind::Binary {
                        op: BinOp::Lt,
                        left: Box::new(ident("i")),
                        right: Box::new(member(ident("s"), "length")),
                    })),
                    right: Box::new(expr(ExprKind::Binary {
                        op: BinOp::GtEq,
                        left: Box::new(call_member(
                            ident("accept"),
                            "indexOf",
                            vec![call_member(ident("s"), "charAt", vec![ident("i")])],
                        )),
                        right: Box::new(int_lit(0)),
                    })),
                }),
                body: vec![stmt(StmtKind::Expr(assign_expr(
                    ident("i"),
                    expr(ExprKind::Binary {
                        op: BinOp::Add,
                        left: Box::new(ident("i")),
                        right: Box::new(int_lit(1)),
                    }),
                )))],
                else_body: None,
            }),
            stmt(StmtKind::Return(Some(ident("i")))),
        ],
    ));

    out.push(function_stmt(
        "__c_strcspn_h",
        vec!["s", "reject"],
        vec![
            var_decl_stmt("i", int_lit(0)),
            stmt(StmtKind::While {
                cond: expr(ExprKind::Binary {
                    op: BinOp::And,
                    left: Box::new(expr(ExprKind::Binary {
                        op: BinOp::Lt,
                        left: Box::new(ident("i")),
                        right: Box::new(member(ident("s"), "length")),
                    })),
                    right: Box::new(expr(ExprKind::Binary {
                        op: BinOp::Lt,
                        left: Box::new(call_member(
                            ident("reject"),
                            "indexOf",
                            vec![call_member(ident("s"), "charAt", vec![ident("i")])],
                        )),
                        right: Box::new(int_lit(0)),
                    })),
                }),
                body: vec![stmt(StmtKind::Expr(assign_expr(
                    ident("i"),
                    expr(ExprKind::Binary {
                        op: BinOp::Add,
                        left: Box::new(ident("i")),
                        right: Box::new(int_lit(1)),
                    }),
                )))],
                else_body: None,
            }),
            stmt(StmtKind::Return(Some(ident("i")))),
        ],
    ));
    out
}

fn bin(op: BinOp, left: vybe_ast::Expression, right: vybe_ast::Expression) -> vybe_ast::Expression {
    expr(ExprKind::Binary { op, left: Box::new(left), right: Box::new(right) })
}

fn is_type(value: vybe_ast::Expression, type_name: &str) -> vybe_ast::Expression {
    expr(ExprKind::IsType { expr: Box::new(value), type_name: type_name.into() })
}

fn ret(value: vybe_ast::Expression) -> Statement {
    stmt(StmtKind::Return(Some(value)))
}

/// `__c_source_byte(src, i)` — byte `i` of any C character source: a string
/// value, a linear address, or a carray pointer over either. Past the end of
/// a string value it reads the terminator.
fn c_source_byte() -> Statement {
    let string_byte = |text: vybe_ast::Expression, index: vybe_ast::Expression| {
        expr(ExprKind::Ternary {
            cond: Box::new(bin(BinOp::Lt, index.clone(), member(text.clone(), "length"))),
            then: Box::new(bin(
                BinOp::BitAnd,
                call_expr(ident("__c_char_code_at"), vec![text, index]),
                int_lit(255),
            )),
            else_: Box::new(int_lit(0)),
        })
    };
    let base = member(ident("src"), "__base");
    let offset = bin(BinOp::Add, member(ident("src"), "__idx"), ident("i"));
    function_stmt(
        "__c_source_byte",
        vec!["src", "i"],
        vec![
            if_stmt(
                is_type(ident("src"), "number"),
                vec![ret(call_expr(
                    ident("__c_ptr_i32_load8_u"),
                    vec![bin(BinOp::Add, ident("src"), ident("i"))],
                ))],
                None,
            ),
            if_stmt(
                is_type(ident("src"), "string"),
                vec![ret(string_byte(ident("src"), ident("i")))],
                None,
            ),
            if_stmt(
                vybe_compiler::primitives::pointers::is_carray_ptr_kind(ident("src")),
                vec![
                    if_stmt(
                        is_type(base.clone(), "string"),
                        vec![ret(string_byte(base.clone(), offset.clone()))],
                        None,
                    ),
                    ret(bin(BinOp::BitAnd, index_expr(base, offset), int_lit(255))),
                ],
                None,
            ),
            ret(int_lit(0)),
        ],
    )
}

/// `__c_memccpy(dst, src, ch, n)` — copy bytes into linear `dst` up to and
/// including the first `ch`; return the address after it, or NULL.
fn c_memccpy() -> Statement {
    let byte = ident("b");
    function_stmt(
        "__c_memccpy",
        vec!["dst", "src", "ch", "n"],
        vec![
            var_decl_stmt("i", int_lit(0)),
            var_decl_stmt("b", int_lit(0)),
            stmt(StmtKind::While {
                cond: bin(BinOp::Lt, ident("i"), ident("n")),
                body: vec![
                    stmt(StmtKind::Expr(assign_expr(
                        byte.clone(),
                        call_expr(ident("__c_source_byte"), vec![ident("src"), ident("i")]),
                    ))),
                    stmt(StmtKind::Expr(call_expr(
                        ident("__c_ptr_i32_store8"),
                        vec![bin(BinOp::Add, ident("dst"), ident("i")), byte.clone()],
                    ))),
                    stmt(StmtKind::Expr(assign_expr(
                        ident("i"),
                        bin(BinOp::Add, ident("i"), int_lit(1)),
                    ))),
                    if_stmt(
                        bin(BinOp::Eq, byte, bin(BinOp::BitAnd, ident("ch"), int_lit(255))),
                        vec![ret(bin(BinOp::Add, ident("dst"), ident("i")))],
                        None,
                    ),
                ],
                else_body: None,
            }),
            ret(null_lit()),
        ],
    )
}

/// `__c_strtok_is_delimiter(c, delim)` — whether byte `c` occurs in the C string
/// `delim` (any character source).
fn c_strtok_is_delimiter() -> Statement {
    let byte = call_expr(ident("__c_source_byte"), vec![ident("delim"), ident("i")]);
    function_stmt(
        "__c_strtok_is_delimiter",
        vec!["c", "delim"],
        vec![
            var_decl_stmt("i", int_lit(0)),
            var_decl_stmt("d", byte.clone()),
            stmt(StmtKind::While {
                cond: bin(BinOp::NotEq, ident("d"), int_lit(0)),
                body: vec![
                    if_stmt(bin(BinOp::Eq, ident("d"), ident("c")), vec![ret(int_lit(1))], None),
                    stmt(StmtKind::Expr(assign_expr(ident("i"), bin(BinOp::Add, ident("i"), int_lit(1))))),
                    stmt(StmtKind::Expr(assign_expr(ident("d"), byte))),
                ],
                else_body: None,
            }),
            ret(int_lit(0)),
        ],
    )
}

/// `__c_strtok(s, delim)` — C `strtok` over a linear string: skip leading
/// delimiters, terminate the token in place, and resume after it on a null
/// `s`. Returns the token's address, or null when none is left.
fn c_strtok() -> Statement {
    let byte = || call_expr(ident("__c_ptr_i32_load8_u"), vec![ident("s")]);
    let is_delim = || call_expr(ident("__c_strtok_is_delimiter"), vec![ident("c"), ident("delim")]);
    let advance = || stmt(StmtKind::Expr(assign_expr(ident("s"), bin(BinOp::Add, ident("s"), int_lit(1)))));
    let save = |value: vybe_ast::Expression| stmt(StmtKind::Expr(assign_expr(ident("__c_strtok_save"), value)));
    let absent = |value: vybe_ast::Expression| {
        bin(BinOp::Or, bin(BinOp::Eq, value.clone(), null_lit()), bin(BinOp::Eq, value, int_lit(0)))
    };
    function_stmt(
        "__c_strtok",
        vec!["s", "delim"],
        vec![
            if_stmt(absent(ident("s")), vec![stmt(StmtKind::Expr(assign_expr(ident("s"), ident("__c_strtok_save"))))], None),
            if_stmt(absent(ident("s")), vec![ret(null_lit())], None),
            var_decl_stmt("c", byte()),
            stmt(StmtKind::While {
                cond: bin(BinOp::And, bin(BinOp::NotEq, ident("c"), int_lit(0)), is_delim()),
                body: vec![advance(), stmt(StmtKind::Expr(assign_expr(ident("c"), byte())))],
                else_body: None,
            }),
            if_stmt(bin(BinOp::Eq, ident("c"), int_lit(0)), vec![save(null_lit()), ret(null_lit())], None),
            var_decl_stmt("token", ident("s")),
            stmt(StmtKind::While {
                cond: bin(BinOp::And, bin(BinOp::NotEq, ident("c"), int_lit(0)),
                    expr(ExprKind::Unary { op: vybe_ast::UnaryOp::Not, expr: Box::new(is_delim()) })),
                body: vec![advance(), stmt(StmtKind::Expr(assign_expr(ident("c"), byte())))],
                else_body: None,
            }),
            if_stmt(bin(BinOp::Eq, ident("c"), int_lit(0)), vec![save(null_lit()), ret(ident("token"))], None),
            stmt(StmtKind::Expr(call_expr(ident("__c_ptr_i32_store8"), vec![ident("s"), int_lit(0)]))),
            save(bin(BinOp::Add, ident("s"), int_lit(1))),
            ret(ident("token")),
        ],
    )
}

/// `__c_strstr_linear(hay, needle)` — C `strstr` over a linear haystack and
/// any character-source needle: the address of the first match, or null.
fn c_strstr_linear() -> Statement {
    let hay = |offset: vybe_ast::Expression| {
        call_expr(ident("__c_ptr_i32_load8_u"), vec![bin(BinOp::Add, ident("hay"), offset)])
    };
    let needle = |index: vybe_ast::Expression| {
        call_expr(ident("__c_source_byte"), vec![ident("needle"), index])
    };
    let set = |name: &str, value: vybe_ast::Expression| stmt(StmtKind::Expr(assign_expr(ident(name), value)));
    let step = |name: &str| set(name, bin(BinOp::Add, ident(name), int_lit(1)));
    function_stmt(
        "__c_strstr_linear",
        vec!["hay", "needle"],
        vec![
            var_decl_stmt("i", int_lit(0)),
            var_decl_stmt("j", int_lit(0)),
            if_stmt(bin(BinOp::Eq, needle(int_lit(0)), int_lit(0)), vec![ret(ident("hay"))], None),
            stmt(StmtKind::While {
                cond: bin(BinOp::NotEq, hay(ident("i")), int_lit(0)),
                body: vec![
                    set("j", int_lit(0)),
                    stmt(StmtKind::While {
                        cond: bin(
                            BinOp::And,
                            bin(BinOp::NotEq, needle(ident("j")), int_lit(0)),
                            bin(BinOp::Eq, hay(bin(BinOp::Add, ident("i"), ident("j"))), needle(ident("j"))),
                        ),
                        body: vec![step("j")],
                        else_body: None,
                    }),
                    if_stmt(
                        bin(BinOp::Eq, needle(ident("j")), int_lit(0)),
                        vec![ret(bin(BinOp::Add, ident("hay"), ident("i")))],
                        None,
                    ),
                    step("i"),
                ],
                else_body: None,
            }),
            ret(null_lit()),
        ],
    )
}

/// `__c_wide_store(dst, units)` — store each code unit of a managed wide
/// array as a 4-byte `wchar_t` into linear `dst`.
fn c_wide_store() -> Statement {
    function_stmt(
        "__c_wide_store",
        vec!["dst", "units"],
        vec![
            var_decl_stmt("i", int_lit(0)),
            stmt(StmtKind::While {
                cond: bin(BinOp::Lt, ident("i"), member(ident("units"), "length")),
                body: vec![
                    stmt(StmtKind::Expr(call_expr(
                        ident("__c_ptr_i32_store"),
                        vec![
                            bin(BinOp::Add, ident("dst"), bin(BinOp::Mul, ident("i"), int_lit(4))),
                            index_expr(ident("units"), ident("i")),
                        ],
                    ))),
                    stmt(StmtKind::Expr(assign_expr(
                        ident("i"),
                        bin(BinOp::Add, ident("i"), int_lit(1)),
                    ))),
                ],
                else_body: None,
            }),
        ],
    )
}

/// `__c_utf8_store(dst, cp)` — encode code point `cp` as UTF-8 into linear
/// `dst` and return the number of bytes written (1-4).
fn c_utf8_store() -> Statement {
    let store = |offset: i64, value: vybe_ast::Expression| {
        stmt(StmtKind::Expr(call_expr(
            ident("__c_ptr_i32_store8"),
            vec![bin(BinOp::Add, ident("dst"), int_lit(offset)), value],
        )))
    };
    let shifted = |bits: i64| bin(BinOp::Shr, ident("cp"), int_lit(bits));
    let continuation = |bits: i64| {
        bin(BinOp::BitOr, int_lit(0x80), bin(BinOp::BitAnd, shifted(bits), int_lit(0x3f)))
    };
    let low = || bin(BinOp::BitOr, int_lit(0x80), bin(BinOp::BitAnd, ident("cp"), int_lit(0x3f)));
    function_stmt(
        "__c_utf8_store",
        vec!["dst", "cp"],
        vec![
            if_stmt(
                bin(BinOp::Lt, ident("cp"), int_lit(0x80)),
                vec![store(0, ident("cp")), ret(int_lit(1))],
                None,
            ),
            if_stmt(
                bin(BinOp::Lt, ident("cp"), int_lit(0x800)),
                vec![
                    store(0, bin(BinOp::BitOr, int_lit(0xc0), shifted(6))),
                    store(1, low()),
                    ret(int_lit(2)),
                ],
                None,
            ),
            if_stmt(
                bin(BinOp::Lt, ident("cp"), int_lit(0x10000)),
                vec![
                    store(0, bin(BinOp::BitOr, int_lit(0xe0), shifted(12))),
                    store(1, continuation(6)),
                    store(2, low()),
                    ret(int_lit(3)),
                ],
                None,
            ),
            store(0, bin(BinOp::BitOr, int_lit(0xf0), shifted(18))),
            store(1, continuation(12)),
            store(2, continuation(6)),
            store(3, low()),
            ret(int_lit(4)),
        ],
    )
}

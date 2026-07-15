//! C string.h runtime helpers (libc surface): the `__c_str*_h` functions backing
//! strcoll / strxfrm / strpbrk / strspn / strcspn. The call-site lowerings live
//! in the walker (they map the C name to these helpers); the helper bodies live
//! here so any libc-targeting front-end can inject them. Injected once into the
//! program prelude.

use crate::emitter::build::*;
use vybe_ast::{BinOp, ExprKind, Statement, StmtKind};

pub fn runtime_helpers() -> Vec<Statement> {
    let mut out: Vec<Statement> = Vec::new();
    out.push(function_stmt(
        "__c_mem_index_of_h",
        vec!["buf", "needle", "n", "reverse"],
        vec![
            var_decl_stmt(
                "limit",
                expr(ExprKind::Ternary {
                    cond: Box::new(expr(ExprKind::Binary {
                        op: BinOp::Lt,
                        left: Box::new(ident("n")),
                        right: Box::new(member(ident("buf"), "length")),
                    })),
                    then: Box::new(ident("n")),
                    else_: Box::new(member(ident("buf"), "length")),
                }),
            ),
            var_decl_stmt(
                "needle_ch",
                call_member(ident("String"), "fromCharCode", vec![ident("needle")]),
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
                    var_decl_stmt("value", index_expr(ident("buf"), ident("i"))),
                    if_stmt(
                        expr(ExprKind::Binary {
                            op: BinOp::Or,
                            left: Box::new(expr(ExprKind::Binary {
                                op: BinOp::Eq,
                                left: Box::new(ident("value")),
                                right: Box::new(ident("needle")),
                            })),
                            right: Box::new(expr(ExprKind::Binary {
                                op: BinOp::Eq,
                                left: Box::new(ident("value")),
                                right: Box::new(ident("needle_ch")),
                            })),
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

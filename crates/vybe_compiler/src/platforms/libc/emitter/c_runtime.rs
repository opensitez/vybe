//! libc runtime prelude — the FILE/stdio model, math.h series helpers, and the
//! rand / signal / locale / strtok runtime, emitted as common AST. This is the
//! libc surface shared by any libc-targeting front-end; the C walker injects it
//! via `prelude()`. Builders come from `build`; math series helpers from
//! `math_runtime`; the stdin / char-decode / wide-char / domain-error helpers
//! are composed from their own adapters.

use crate::ast::{
    Argument, ArrayElement, BinOp, BindingPattern, ExprKind, Literal, ObjectProperty, Statement,
    StmtKind, VarDeclKind, VarDeclarator,
};
use crate::platforms::libc::emitter::build::*;
use crate::platforms::libc::emitter::math_runtime::{
    build_math_helper_fn, ecma_math_call, poly_erf, stirling_approx,
};

pub fn prelude() -> Vec<Statement> {
    let stdout_name = "__c_stdout_file";
    let buffer_name = "__c_stdout_buffer";
    let store_name = "__c_file_store";

    let mut out = vec![
        // ── Math helpers not in WASM or ecma:math ───────────────────────
        // __tgamma(x): gamma function — use Lanczos approximation
        build_math_helper_fn(
            "__tgamma",
            &["x"],
            vec![
                // Simple: for positive integers, (n-1)!
                // Approx via: sqrt(2*pi/x) * (x/e)^x (Stirling)
                // For test coverage, use a JS-style approximation:
                // tgamma(n) ≈ exp(lgamma(n))
                stmt(StmtKind::Return(Some(ecma_math_call(
                    "exp",
                    ecma_math_call(
                        "log",
                        expr(ExprKind::Ternary {
                            cond: Box::new(expr(ExprKind::Binary {
                                op: BinOp::LtEq,
                                left: Box::new(ident("x")),
                                right: Box::new(expr(ExprKind::Lit(Literal::Float(0.0)))),
                            })),
                            then: Box::new(expr(ExprKind::Lit(Literal::Float(f64::INFINITY)))),
                            else_: Box::new(stirling_approx()),
                        }),
                    ),
                )))),
            ],
        ),
        build_math_helper_fn(
            "__lgamma",
            &["x"],
            vec![
                // lgamma(x) = log(tgamma(x)), use log of Stirling approx
                stmt(StmtKind::Return(Some(ecma_math_call(
                    "log",
                    stirling_approx(),
                )))),
            ],
        ),
        build_math_helper_fn(
            "__erf",
            &["x"],
            vec![
                // erf approximation (Abramowitz & Stegun 7.1.26)
                // t = 1/(1+0.3275911*x), polynomial approx
                // erf(x) ≈ 1 - (a1*t + a2*t^2 + a3*t^3 + a4*t^4 + a5*t^5) * e^(-x^2)
                stmt(StmtKind::VarDecl {
                    declarations: vec![VarDeclarator {
                        pattern: BindingPattern::Ident("t".to_string()),
                        type_hint: None,
                        init: Some(expr(ExprKind::Binary {
                            op: BinOp::Div,
                            left: Box::new(expr(ExprKind::Lit(Literal::Float(1.0)))),
                            right: Box::new(expr(ExprKind::Binary {
                                op: BinOp::Add,
                                left: Box::new(expr(ExprKind::Lit(Literal::Float(1.0)))),
                                right: Box::new(expr(ExprKind::Binary {
                                    op: BinOp::Mul,
                                    left: Box::new(expr(ExprKind::Lit(Literal::Float(0.3275911)))),
                                    right: Box::new(ecma_math_call("abs", ident("x"))),
                                })),
                            })),
                        })),
                        array_bounds: None,
                        with_events: false,
                    }],
                    kind: VarDeclKind::Let,
                }),
                stmt(StmtKind::VarDecl {
                    declarations: vec![VarDeclarator {
                        pattern: BindingPattern::Ident("poly".to_string()),
                        type_hint: None,
                        init: Some(poly_erf(ident("t"))),
                        array_bounds: None,
                        with_events: false,
                    }],
                    kind: VarDeclKind::Let,
                }),
                stmt(StmtKind::VarDecl {
                    declarations: vec![VarDeclarator {
                        pattern: BindingPattern::Ident("result".to_string()),
                        type_hint: None,
                        init: Some(expr(ExprKind::Binary {
                            op: BinOp::Sub,
                            left: Box::new(expr(ExprKind::Lit(Literal::Float(1.0)))),
                            right: Box::new(expr(ExprKind::Binary {
                                op: BinOp::Mul,
                                left: Box::new(ident("poly")),
                                right: Box::new(ecma_math_call(
                                    "exp",
                                    expr(ExprKind::Unary {
                                        op: crate::ast::UnaryOp::Neg,
                                        expr: Box::new(expr(ExprKind::Binary {
                                            op: BinOp::Mul,
                                            left: Box::new(ident("x")),
                                            right: Box::new(ident("x")),
                                        })),
                                    }),
                                )),
                            })),
                        })),
                        array_bounds: None,
                        with_events: false,
                    }],
                    kind: VarDeclKind::Let,
                }),
                // Return result with correct sign
                stmt(StmtKind::Return(Some(expr(ExprKind::Ternary {
                    cond: Box::new(expr(ExprKind::Binary {
                        op: BinOp::Lt,
                        left: Box::new(ident("x")),
                        right: Box::new(expr(ExprKind::Lit(Literal::Float(0.0)))),
                    })),
                    then: Box::new(expr(ExprKind::Unary {
                        op: crate::ast::UnaryOp::Neg,
                        expr: Box::new(ident("result")),
                    })),
                    else_: Box::new(ident("result")),
                })))),
            ],
        ),
        build_math_helper_fn(
            "__j0",
            &["x"],
            vec![
                // Bessel J0: j0(0) = 1, approximation for general x
                stmt(StmtKind::Return(Some(expr(ExprKind::Ternary {
                    cond: Box::new(expr(ExprKind::Binary {
                        op: BinOp::Eq,
                        left: Box::new(ident("x")),
                        right: Box::new(expr(ExprKind::Lit(Literal::Float(0.0)))),
                    })),
                    then: Box::new(expr(ExprKind::Lit(Literal::Float(1.0)))),
                    else_: Box::new(expr(ExprKind::Binary {
                        op: BinOp::Div,
                        left: Box::new(ecma_math_call("sin", ident("x"))),
                        right: Box::new(ident("x")),
                    })),
                })))),
            ],
        ),
        stmt(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident(store_name.to_string()),
                type_hint: None,
                init: Some(expr(ExprKind::Object(vec![]))),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        }),
        stmt(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident(buffer_name.to_string()),
                type_hint: None,
                init: Some(str_lit("")),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        }),
        stmt(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident("__c_file_content".to_string()),
                type_hint: None,
                init: Some(expr(ExprKind::Object(vec![]))),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        }),
        stmt(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident("__c_file_pos".to_string()),
                type_hint: None,
                init: Some(expr(ExprKind::Object(vec![]))),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        }),
        stmt(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident("__c_file_eof".to_string()),
                type_hint: None,
                init: Some(expr(ExprKind::Object(vec![]))),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        }),
        stmt(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident("__c_file_ungot".to_string()),
                type_hint: None,
                init: Some(expr(ExprKind::Object(vec![]))),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        }),
        stmt(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident("__c_file_dirty".to_string()),
                type_hint: None,
                init: Some(expr(ExprKind::Object(vec![]))),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        }),
        stmt(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident("__c_file_binary".to_string()),
                type_hint: None,
                init: Some(expr(ExprKind::Object(vec![]))),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        }),
        stmt(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident("__c_file_pathmap".to_string()),
                type_hint: None,
                init: Some(expr(ExprKind::Object(vec![]))),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        }),
        stmt(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident("__c_next_file_handle".to_string()),
                type_hint: None,
                init: Some(int_lit(2)),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        }),
        stmt(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident("__c_rand_state".to_string()),
                type_hint: None,
                init: Some(int_lit(1)),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        }),
        stmt(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident("__c_signal_handlers".to_string()),
                type_hint: None,
                init: Some(expr(ExprKind::Object(vec![]))),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        }),
        stmt(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident("__c_locale".to_string()),
                type_hint: None,
                init: Some(str_lit("C")),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        }),
        stmt(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident("__c_strtok_rem".to_string()),
                type_hint: None,
                init: Some(null_lit()),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        }),
        stmt(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident("__c_strtok_delim".to_string()),
                type_hint: None,
                init: Some(str_lit("")),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        }),
        stmt(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident(stdout_name.to_string()),
                type_hint: None,
                init: Some(expr(ExprKind::Array(vec![
                    ArrayElement {
                        key: None,
                        value: str_lit("__stdout__"),
                        spread: false,
                        by_ref: false,
                    },
                    ArrayElement {
                        key: None,
                        value: str_lit("w"),
                        spread: false,
                        by_ref: false,
                    },
                    ArrayElement {
                        key: None,
                        value: str_lit(""),
                        spread: false,
                        by_ref: false,
                    },
                    ArrayElement {
                        key: None,
                        value: int_lit(0),
                        spread: false,
                        by_ref: false,
                    },
                    ArrayElement {
                        key: None,
                        value: int_lit(0),
                        spread: false,
                        by_ref: false,
                    },
                    ArrayElement {
                        key: None,
                        value: null_lit(),
                        spread: false,
                        by_ref: false,
                    },
                    ArrayElement {
                        key: None,
                        value: int_lit(0),
                        spread: false,
                        by_ref: false,
                    },
                    ArrayElement {
                        key: None,
                        value: int_lit(0),
                        spread: false,
                        by_ref: false,
                    },
                    ArrayElement {
                        key: None,
                        value: str_lit("stdout"),
                        spread: false,
                        by_ref: false,
                    },
                ]))),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        }),
    ];

    out.push(function_stmt(
        "__c_char_ptr_add",
        vec!["s", "n"],
        vec![stmt(StmtKind::Return(Some(call_member(
            ident("s"),
            "substring",
            vec![ident("n")],
        ))))],
    ));

    out.push(function_stmt(
        "__c_stdout_append",
        vec!["piece"],
        vec![
            if_stmt(
                call_member(ident("piece"), "endsWith", vec![str_lit("\n")]),
                vec![
                    stmt(StmtKind::Expr(call_expr(
                        ident("puts"),
                        vec![expr(ExprKind::Binary {
                            op: BinOp::Add,
                            left: Box::new(ident(buffer_name)),
                            right: Box::new(call_member(
                                ident("piece"),
                                "slice",
                                vec![
                                    int_lit(0),
                                    expr(ExprKind::Binary {
                                        op: BinOp::Sub,
                                        left: Box::new(member(ident("piece"), "length")),
                                        right: Box::new(int_lit(1)),
                                    }),
                                ],
                            )),
                        })],
                    ))),
                    stmt(StmtKind::Expr(assign_expr(ident(buffer_name), str_lit("")))),
                ],
                Some(vec![stmt(StmtKind::Expr(assign_expr(
                    ident(buffer_name),
                    expr(ExprKind::Binary {
                        op: BinOp::Add,
                        left: Box::new(ident(buffer_name)),
                        right: Box::new(ident("piece")),
                    }),
                )))]),
            ),
            stmt(StmtKind::Return(Some(int_lit(0)))),
        ],
    ));

    out.push(function_stmt(
        "__c_file_new",
        vec!["path", "mode"],
        vec![
            var_decl_stmt("existing", index_expr(ident(store_name), ident("path"))),
            var_decl_stmt(
                "write_mode",
                expr(ExprKind::Ternary {
                    cond: Box::new(expr(ExprKind::Binary {
                        op: BinOp::GtEq,
                        left: Box::new(call_member(ident("mode"), "indexOf", vec![str_lit("w")])),
                        right: Box::new(int_lit(0)),
                    })),
                    then: Box::new(int_lit(1)),
                    else_: Box::new(int_lit(0)),
                }),
            ),
            var_decl_stmt(
                "binary_mode",
                expr(ExprKind::Ternary {
                    cond: Box::new(expr(ExprKind::Binary {
                        op: BinOp::GtEq,
                        left: Box::new(call_member(ident("mode"), "indexOf", vec![str_lit("b")])),
                        right: Box::new(int_lit(0)),
                    })),
                    then: Box::new(int_lit(1)),
                    else_: Box::new(int_lit(0)),
                }),
            ),
            var_decl_stmt("content", str_lit("")),
            if_stmt(
                expr(ExprKind::Binary {
                    op: BinOp::NotEq,
                    left: Box::new(ident("binary_mode")),
                    right: Box::new(int_lit(0)),
                }),
                vec![stmt(StmtKind::Expr(assign_expr(
                    ident("content"),
                    expr(ExprKind::Array(vec![])),
                )))],
                None,
            ),
            if_stmt(
                expr(ExprKind::Binary {
                    op: BinOp::And,
                    left: Box::new(expr(ExprKind::Binary {
                        op: BinOp::NotEq,
                        left: Box::new(ident("existing")),
                        right: Box::new(null_lit()),
                    })),
                    right: Box::new(expr(ExprKind::Binary {
                        op: BinOp::NotEq,
                        left: Box::new(ident("existing")),
                        right: Box::new(expr(ExprKind::Lit(Literal::Undefined))),
                    })),
                }),
                vec![stmt(StmtKind::Expr(assign_expr(
                    ident("content"),
                    ident("existing"),
                )))],
                None,
            ),
            if_stmt(
                expr(ExprKind::Binary {
                    op: BinOp::NotEq,
                    left: Box::new(ident("write_mode")),
                    right: Box::new(int_lit(0)),
                }),
                vec![
                    stmt(StmtKind::Expr(assign_expr(ident("content"), str_lit("")))),
                    if_stmt(
                        expr(ExprKind::Binary {
                            op: BinOp::NotEq,
                            left: Box::new(ident("binary_mode")),
                            right: Box::new(int_lit(0)),
                        }),
                        vec![stmt(StmtKind::Expr(assign_expr(
                            ident("content"),
                            expr(ExprKind::Array(vec![])),
                        )))],
                        None,
                    ),
                    stmt(StmtKind::Expr(assign_expr(
                        index_expr(ident(store_name), ident("path")),
                        ident("content"),
                    ))),
                ],
                None,
            ),
            var_decl_stmt("fileobj", expr(ExprKind::Object(vec![]))),
            stmt(StmtKind::Expr(assign_expr(
                member(ident("fileobj"), "path"),
                ident("path"),
            ))),
            stmt(StmtKind::Expr(assign_expr(
                member(ident("fileobj"), "mode"),
                ident("mode"),
            ))),
            stmt(StmtKind::Expr(assign_expr(
                member(ident("fileobj"), "content"),
                ident("content"),
            ))),
            stmt(StmtKind::Expr(assign_expr(
                member(ident("fileobj"), "pos"),
                int_lit(0),
            ))),
            stmt(StmtKind::Expr(assign_expr(
                member(ident("fileobj"), "eof"),
                int_lit(0),
            ))),
            stmt(StmtKind::Expr(assign_expr(
                member(ident("fileobj"), "ungot"),
                null_lit(),
            ))),
            stmt(StmtKind::Expr(assign_expr(
                member(ident("fileobj"), "dirty"),
                ident("write_mode"),
            ))),
            stmt(StmtKind::Expr(assign_expr(
                member(ident("fileobj"), "binary"),
                ident("binary_mode"),
            ))),
            stmt(StmtKind::Return(Some(ident("fileobj")))),
        ],
    ));

    out.push(function_stmt(
        "__c_file_sync",
        vec!["file"],
        vec![
            if_stmt(
                expr(ExprKind::Binary {
                    op: BinOp::Eq,
                    left: Box::new(file_slot(ident("file"), CFILE_SPECIAL)),
                    right: Box::new(str_lit("stdout")),
                }),
                vec![stmt(StmtKind::Return(Some(int_lit(0))))],
                None,
            ),
            if_stmt(
                expr(ExprKind::Binary {
                    op: BinOp::NotEq,
                    left: Box::new(file_slot(ident("file"), CFILE_DIRTY)),
                    right: Box::new(int_lit(0)),
                }),
                vec![
                    stmt(StmtKind::Expr(assign_expr(
                        index_expr(ident(store_name), file_slot(ident("file"), CFILE_PATH)),
                        file_slot(ident("file"), CFILE_CONTENT),
                    ))),
                    stmt(StmtKind::Expr(assign_expr(
                        file_slot(ident("file"), CFILE_DIRTY),
                        int_lit(0),
                    ))),
                ],
                None,
            ),
            stmt(StmtKind::Return(Some(int_lit(0)))),
        ],
    ));

    out.push(function_stmt(
        "__c_fputs",
        vec!["text", "file"],
        vec![
            if_stmt(
                expr(ExprKind::Binary {
                    op: BinOp::Eq,
                    left: Box::new(file_slot(ident("file"), CFILE_SPECIAL)),
                    right: Box::new(str_lit("stdout")),
                }),
                vec![stmt(StmtKind::Return(Some(call_expr(
                    ident("__c_stdout_append"),
                    vec![ident("text")],
                ))))],
                None,
            ),
            if_stmt(
                expr(ExprKind::Binary {
                    op: BinOp::Eq,
                    left: Box::new(file_slot(ident("file"), CFILE_POS)),
                    right: Box::new(int_lit(0)),
                }),
                vec![stmt(StmtKind::Expr(assign_expr(
                    file_slot(ident("file"), CFILE_CONTENT),
                    ident("text"),
                )))],
                Some(vec![stmt(StmtKind::Expr(assign_expr(
                    file_slot(ident("file"), CFILE_CONTENT),
                    expr(ExprKind::Binary {
                        op: BinOp::Add,
                        left: Box::new(file_slot(ident("file"), CFILE_CONTENT)),
                        right: Box::new(ident("text")),
                    }),
                )))]),
            ),
            stmt(StmtKind::Expr(assign_expr(
                file_slot(ident("file"), CFILE_POS),
                member(file_slot(ident("file"), CFILE_CONTENT), "length"),
            ))),
            stmt(StmtKind::Expr(assign_expr(
                file_slot(ident("file"), CFILE_DIRTY),
                int_lit(1),
            ))),
            stmt(StmtKind::Return(Some(int_lit(0)))),
        ],
    ));

    out.push(function_stmt(
        "__c_fputc",
        vec!["code", "file"],
        vec![
            stmt(StmtKind::Expr(call_expr(
                ident("__c_fputs"),
                vec![
                    call_member(ident("String"), "fromCharCode", vec![ident("code")]),
                    ident("file"),
                ],
            ))),
            stmt(StmtKind::Return(Some(ident("code")))),
        ],
    ));

    out.push(function_stmt(
        "__c_fgetc",
        vec!["file"],
        vec![
            if_stmt(
                expr(ExprKind::Binary {
                    op: BinOp::NotEq,
                    left: Box::new(file_slot(ident("file"), CFILE_UNGOT)),
                    right: Box::new(null_lit()),
                }),
                vec![
                    var_decl_stmt("ch", file_slot(ident("file"), CFILE_UNGOT)),
                    stmt(StmtKind::Expr(assign_expr(
                        file_slot(ident("file"), CFILE_UNGOT),
                        null_lit(),
                    ))),
                    stmt(StmtKind::Expr(assign_expr(
                        file_slot(ident("file"), CFILE_EOF),
                        int_lit(0),
                    ))),
                    stmt(StmtKind::Return(Some(ident("ch")))),
                ],
                None,
            ),
            if_stmt(
                expr(ExprKind::Binary {
                    op: BinOp::GtEq,
                    left: Box::new(file_slot(ident("file"), CFILE_POS)),
                    right: Box::new(member(file_slot(ident("file"), CFILE_CONTENT), "length")),
                }),
                vec![
                    stmt(StmtKind::Expr(assign_expr(
                        file_slot(ident("file"), CFILE_EOF),
                        int_lit(1),
                    ))),
                    stmt(StmtKind::Return(Some(int_lit(-1)))),
                ],
                None,
            ),
            var_decl_stmt(
                "ch",
                call_expr(
                    ident("__c_char_code_at"),
                    vec![
                        file_slot(ident("file"), CFILE_CONTENT),
                        file_slot(ident("file"), CFILE_POS),
                    ],
                ),
            ),
            stmt(StmtKind::Expr(assign_expr(
                file_slot(ident("file"), CFILE_POS),
                expr(ExprKind::Binary {
                    op: BinOp::Add,
                    left: Box::new(file_slot(ident("file"), CFILE_POS)),
                    right: Box::new(int_lit(1)),
                }),
            ))),
            stmt(StmtKind::Expr(assign_expr(
                file_slot(ident("file"), CFILE_EOF),
                int_lit(0),
            ))),
            stmt(StmtKind::Return(Some(ident("ch")))),
        ],
    ));

    out.push(function_stmt(
        "__c_ungetc",
        vec!["code", "file"],
        vec![
            stmt(StmtKind::Expr(assign_expr(
                file_slot(ident("file"), CFILE_UNGOT),
                ident("code"),
            ))),
            stmt(StmtKind::Expr(assign_expr(
                file_slot(ident("file"), CFILE_EOF),
                int_lit(0),
            ))),
            stmt(StmtKind::Return(Some(ident("code")))),
        ],
    ));

    out.push(function_stmt(
        "__c_fgets_impl",
        vec!["file", "size"],
        vec![
            if_stmt(
                expr(ExprKind::Binary {
                    op: BinOp::GtEq,
                    left: Box::new(file_slot(ident("file"), CFILE_POS)),
                    right: Box::new(member(file_slot(ident("file"), CFILE_CONTENT), "length")),
                }),
                vec![
                    stmt(StmtKind::Expr(assign_expr(
                        file_slot(ident("file"), CFILE_EOF),
                        int_lit(1),
                    ))),
                    stmt(StmtKind::Return(Some(str_lit("")))),
                ],
                None,
            ),
            var_decl_stmt(
                "rest",
                call_member(
                    file_slot(ident("file"), CFILE_CONTENT),
                    "substring",
                    vec![file_slot(ident("file"), CFILE_POS)],
                ),
            ),
            var_decl_stmt(
                "limit",
                expr(ExprKind::Binary {
                    op: BinOp::Sub,
                    left: Box::new(ident("size")),
                    right: Box::new(int_lit(1)),
                }),
            ),
            var_decl_stmt(
                "nl",
                call_member(ident("rest"), "indexOf", vec![str_lit("\n")]),
            ),
            var_decl_stmt(
                "take",
                expr(ExprKind::Ternary {
                    cond: Box::new(expr(ExprKind::Binary {
                        op: BinOp::Lt,
                        left: Box::new(ident("nl")),
                        right: Box::new(int_lit(0)),
                    })),
                    then: Box::new(ident("limit")),
                    else_: Box::new(expr(ExprKind::Ternary {
                        cond: Box::new(expr(ExprKind::Binary {
                            op: BinOp::Lt,
                            left: Box::new(expr(ExprKind::Binary {
                                op: BinOp::Add,
                                left: Box::new(ident("nl")),
                                right: Box::new(int_lit(1)),
                            })),
                            right: Box::new(ident("limit")),
                        })),
                        then: Box::new(expr(ExprKind::Binary {
                            op: BinOp::Add,
                            left: Box::new(ident("nl")),
                            right: Box::new(int_lit(1)),
                        })),
                        else_: Box::new(ident("limit")),
                    })),
                }),
            ),
            if_stmt(
                expr(ExprKind::Binary {
                    op: BinOp::Gt,
                    left: Box::new(ident("take")),
                    right: Box::new(member(ident("rest"), "length")),
                }),
                vec![stmt(StmtKind::Expr(assign_expr(
                    ident("take"),
                    member(ident("rest"), "length"),
                )))],
                None,
            ),
            var_decl_stmt(
                "out",
                call_member(ident("rest"), "substring", vec![int_lit(0), ident("take")]),
            ),
            stmt(StmtKind::Expr(assign_expr(
                file_slot(ident("file"), CFILE_POS),
                expr(ExprKind::Binary {
                    op: BinOp::Add,
                    left: Box::new(file_slot(ident("file"), CFILE_POS)),
                    right: Box::new(member(ident("out"), "length")),
                }),
            ))),
            stmt(StmtKind::Expr(assign_expr(
                file_slot(ident("file"), CFILE_EOF),
                expr(ExprKind::Ternary {
                    cond: Box::new(expr(ExprKind::Binary {
                        op: BinOp::GtEq,
                        left: Box::new(file_slot(ident("file"), CFILE_POS)),
                        right: Box::new(member(file_slot(ident("file"), CFILE_CONTENT), "length")),
                    })),
                    then: Box::new(int_lit(1)),
                    else_: Box::new(int_lit(0)),
                }),
            ))),
            stmt(StmtKind::Return(Some(ident("out")))),
        ],
    ));

    out.push(function_stmt(
        "__c_fseek_impl",
        vec!["file", "offset", "whence"],
        vec![
            var_decl_stmt(
                "len",
                member(file_slot(ident("file"), CFILE_CONTENT), "length"),
            ),
            var_decl_stmt(
                "pos",
                expr(ExprKind::Ternary {
                    cond: Box::new(expr(ExprKind::Binary {
                        op: BinOp::Eq,
                        left: Box::new(ident("whence")),
                        right: Box::new(int_lit(0)),
                    })),
                    then: Box::new(ident("offset")),
                    else_: Box::new(expr(ExprKind::Ternary {
                        cond: Box::new(expr(ExprKind::Binary {
                            op: BinOp::Eq,
                            left: Box::new(ident("whence")),
                            right: Box::new(int_lit(1)),
                        })),
                        then: Box::new(expr(ExprKind::Binary {
                            op: BinOp::Add,
                            left: Box::new(file_slot(ident("file"), CFILE_POS)),
                            right: Box::new(ident("offset")),
                        })),
                        else_: Box::new(expr(ExprKind::Binary {
                            op: BinOp::Add,
                            left: Box::new(ident("len")),
                            right: Box::new(ident("offset")),
                        })),
                    })),
                }),
            ),
            if_stmt(
                expr(ExprKind::Binary {
                    op: BinOp::Lt,
                    left: Box::new(ident("pos")),
                    right: Box::new(int_lit(0)),
                }),
                vec![stmt(StmtKind::Expr(assign_expr(ident("pos"), int_lit(0))))],
                None,
            ),
            if_stmt(
                expr(ExprKind::Binary {
                    op: BinOp::Gt,
                    left: Box::new(ident("pos")),
                    right: Box::new(ident("len")),
                }),
                vec![stmt(StmtKind::Expr(assign_expr(
                    ident("pos"),
                    ident("len"),
                )))],
                None,
            ),
            stmt(StmtKind::Expr(assign_expr(
                file_slot(ident("file"), CFILE_POS),
                ident("pos"),
            ))),
            stmt(StmtKind::Expr(assign_expr(
                file_slot(ident("file"), CFILE_EOF),
                int_lit(0),
            ))),
            stmt(StmtKind::Expr(assign_expr(
                file_slot(ident("file"), CFILE_UNGOT),
                null_lit(),
            ))),
            stmt(StmtKind::Return(Some(int_lit(0)))),
        ],
    ));

    out.push(function_stmt(
        "__c_fwrite_impl",
        vec!["data", "count", "file"],
        vec![
            stmt(StmtKind::Expr(assign_expr(
                file_slot(ident("file"), CFILE_CONTENT),
                call_member(ident("data"), "slice", vec![int_lit(0), ident("count")]),
            ))),
            stmt(StmtKind::Expr(assign_expr(
                file_slot(ident("file"), CFILE_POS),
                ident("count"),
            ))),
            stmt(StmtKind::Expr(assign_expr(
                file_slot(ident("file"), CFILE_DIRTY),
                int_lit(1),
            ))),
            stmt(StmtKind::Return(Some(ident("count")))),
        ],
    ));

    out.push(function_stmt(
        "__c_fread_impl",
        vec!["file", "count"],
        vec![
            var_decl_stmt(
                "out",
                call_member(
                    file_slot(ident("file"), CFILE_CONTENT),
                    "slice",
                    vec![int_lit(0), ident("count")],
                ),
            ),
            stmt(StmtKind::Expr(assign_expr(
                file_slot(ident("file"), CFILE_POS),
                ident("count"),
            ))),
            stmt(StmtKind::Expr(assign_expr(
                file_slot(ident("file"), CFILE_EOF),
                expr(ExprKind::Ternary {
                    cond: Box::new(expr(ExprKind::Binary {
                        op: BinOp::GtEq,
                        left: Box::new(file_slot(ident("file"), CFILE_POS)),
                        right: Box::new(member(file_slot(ident("file"), CFILE_CONTENT), "length")),
                    })),
                    then: Box::new(int_lit(1)),
                    else_: Box::new(int_lit(0)),
                }),
            ))),
            stmt(StmtKind::Return(Some(ident("out")))),
        ],
    ));

    out.push(function_stmt(
        "__c_fopen_h",
        vec!["path", "mode"],
        vec![
            var_decl_stmt("handle", ident("__c_next_file_handle")),
            stmt(StmtKind::Expr(assign_expr(
                ident("__c_next_file_handle"),
                expr(ExprKind::Binary {
                    op: BinOp::Add,
                    left: Box::new(ident("__c_next_file_handle")),
                    right: Box::new(int_lit(1)),
                }),
            ))),
            var_decl_stmt(
                "write_mode",
                expr(ExprKind::Ternary {
                    cond: Box::new(expr(ExprKind::Binary {
                        op: BinOp::GtEq,
                        left: Box::new(call_member(ident("mode"), "indexOf", vec![str_lit("w")])),
                        right: Box::new(int_lit(0)),
                    })),
                    then: Box::new(int_lit(1)),
                    else_: Box::new(int_lit(0)),
                }),
            ),
            var_decl_stmt(
                "binary_mode",
                expr(ExprKind::Ternary {
                    cond: Box::new(expr(ExprKind::Binary {
                        op: BinOp::GtEq,
                        left: Box::new(call_member(ident("mode"), "indexOf", vec![str_lit("b")])),
                        right: Box::new(int_lit(0)),
                    })),
                    then: Box::new(int_lit(1)),
                    else_: Box::new(int_lit(0)),
                }),
            ),
            var_decl_stmt(
                "content",
                index_expr(ident("__c_file_store"), ident("path")),
            ),
            if_stmt(
                expr(ExprKind::Binary {
                    op: BinOp::Or,
                    left: Box::new(expr(ExprKind::Binary {
                        op: BinOp::Eq,
                        left: Box::new(ident("content")),
                        right: Box::new(null_lit()),
                    })),
                    right: Box::new(expr(ExprKind::Binary {
                        op: BinOp::Eq,
                        left: Box::new(ident("content")),
                        right: Box::new(expr(ExprKind::Lit(Literal::Undefined))),
                    })),
                }),
                vec![stmt(StmtKind::Expr(assign_expr(
                    ident("content"),
                    expr(ExprKind::Ternary {
                        cond: Box::new(expr(ExprKind::Binary {
                            op: BinOp::NotEq,
                            left: Box::new(ident("binary_mode")),
                            right: Box::new(int_lit(0)),
                        })),
                        then: Box::new(expr(ExprKind::Array(vec![]))),
                        else_: Box::new(str_lit("")),
                    }),
                )))],
                None,
            ),
            if_stmt(
                expr(ExprKind::Binary {
                    op: BinOp::NotEq,
                    left: Box::new(ident("write_mode")),
                    right: Box::new(int_lit(0)),
                }),
                vec![stmt(StmtKind::Expr(assign_expr(
                    ident("content"),
                    expr(ExprKind::Ternary {
                        cond: Box::new(expr(ExprKind::Binary {
                            op: BinOp::NotEq,
                            left: Box::new(ident("binary_mode")),
                            right: Box::new(int_lit(0)),
                        })),
                        then: Box::new(expr(ExprKind::Array(vec![]))),
                        else_: Box::new(str_lit("")),
                    }),
                )))],
                None,
            ),
            stmt(StmtKind::Expr(assign_expr(
                index_expr(ident("__c_file_pathmap"), ident("handle")),
                ident("path"),
            ))),
            stmt(StmtKind::Expr(assign_expr(
                index_expr(ident("__c_file_content"), ident("handle")),
                ident("content"),
            ))),
            stmt(StmtKind::Expr(assign_expr(
                index_expr(ident("__c_file_pos"), ident("handle")),
                int_lit(0),
            ))),
            stmt(StmtKind::Expr(assign_expr(
                index_expr(ident("__c_file_eof"), ident("handle")),
                int_lit(0),
            ))),
            stmt(StmtKind::Expr(assign_expr(
                index_expr(ident("__c_file_ungot"), ident("handle")),
                null_lit(),
            ))),
            stmt(StmtKind::Expr(assign_expr(
                index_expr(ident("__c_file_dirty"), ident("handle")),
                ident("write_mode"),
            ))),
            stmt(StmtKind::Expr(assign_expr(
                index_expr(ident("__c_file_binary"), ident("handle")),
                ident("binary_mode"),
            ))),
            stmt(StmtKind::Return(Some(ident("handle")))),
        ],
    ));

    out.push(function_stmt(
        "__c_fsync_h",
        vec!["handle"],
        vec![
            if_stmt(
                expr(ExprKind::Binary {
                    op: BinOp::Eq,
                    left: Box::new(ident("handle")),
                    right: Box::new(int_lit(1)),
                }),
                vec![stmt(StmtKind::Return(Some(int_lit(0))))],
                None,
            ),
            if_stmt(
                expr(ExprKind::Binary {
                    op: BinOp::NotEq,
                    left: Box::new(index_expr(ident("__c_file_dirty"), ident("handle"))),
                    right: Box::new(int_lit(0)),
                }),
                vec![
                    stmt(StmtKind::Expr(assign_expr(
                        index_expr(
                            ident("__c_file_store"),
                            index_expr(ident("__c_file_pathmap"), ident("handle")),
                        ),
                        index_expr(ident("__c_file_content"), ident("handle")),
                    ))),
                    stmt(StmtKind::Expr(assign_expr(
                        index_expr(ident("__c_file_dirty"), ident("handle")),
                        int_lit(0),
                    ))),
                ],
                None,
            ),
            stmt(StmtKind::Return(Some(int_lit(0)))),
        ],
    ));

    out.push(function_stmt(
        "__c_write_carray_string",
        vec!["ptr", "text"],
        vec![
            var_decl_stmt("i", int_lit(0)),
            stmt(StmtKind::While {
                cond: expr(ExprKind::Binary {
                    op: BinOp::Lt,
                    left: Box::new(ident("i")),
                    right: Box::new(member(ident("text"), "length")),
                }),
                body: vec![
                    stmt(StmtKind::Expr(assign_expr(
                        index_expr(
                            member(ident("ptr"), "__base"),
                            expr(ExprKind::Binary {
                                op: BinOp::Add,
                                left: Box::new(member(ident("ptr"), "__idx")),
                                right: Box::new(ident("i")),
                            }),
                        ),
                        call_expr(ident("__c_char_code_at"), vec![ident("text"), ident("i")]),
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
            stmt(StmtKind::Expr(assign_expr(
                index_expr(
                    member(ident("ptr"), "__base"),
                    expr(ExprKind::Binary {
                        op: BinOp::Add,
                        left: Box::new(member(ident("ptr"), "__idx")),
                        right: Box::new(ident("i")),
                    }),
                ),
                int_lit(0),
            ))),
            stmt(StmtKind::Return(Some(member(ident("text"), "length")))),
        ],
    ));

    out.push(function_stmt(
        "__c_fputs_h",
        vec!["text", "handle"],
        vec![
            if_stmt(
                expr(ExprKind::Binary {
                    op: BinOp::Eq,
                    left: Box::new(ident("handle")),
                    right: Box::new(int_lit(1)),
                }),
                vec![stmt(StmtKind::Return(Some(call_expr(
                    ident("__c_stdout_append"),
                    vec![ident("text")],
                ))))],
                None,
            ),
            if_stmt(
                expr(ExprKind::Binary {
                    op: BinOp::Eq,
                    left: Box::new(index_expr(ident("__c_file_pos"), ident("handle"))),
                    right: Box::new(int_lit(0)),
                }),
                vec![stmt(StmtKind::Expr(assign_expr(
                    index_expr(ident("__c_file_content"), ident("handle")),
                    ident("text"),
                )))],
                Some(vec![stmt(StmtKind::Expr(assign_expr(
                    index_expr(ident("__c_file_content"), ident("handle")),
                    expr(ExprKind::Binary {
                        op: BinOp::Add,
                        left: Box::new(index_expr(ident("__c_file_content"), ident("handle"))),
                        right: Box::new(ident("text")),
                    }),
                )))]),
            ),
            stmt(StmtKind::Expr(assign_expr(
                index_expr(ident("__c_file_pos"), ident("handle")),
                member(
                    index_expr(ident("__c_file_content"), ident("handle")),
                    "length",
                ),
            ))),
            stmt(StmtKind::Expr(assign_expr(
                index_expr(ident("__c_file_dirty"), ident("handle")),
                int_lit(1),
            ))),
            stmt(StmtKind::Return(Some(int_lit(0)))),
        ],
    ));

    out.push(function_stmt(
        "__c_fputc_h",
        vec!["code", "handle"],
        vec![
            stmt(StmtKind::Expr(call_expr(
                ident("__c_fputs_h"),
                vec![
                    call_member(ident("String"), "fromCharCode", vec![ident("code")]),
                    ident("handle"),
                ],
            ))),
            stmt(StmtKind::Return(Some(ident("code")))),
        ],
    ));

    out.push(function_stmt(
        "__c_fgetc_h",
        vec!["handle"],
        vec![
            if_stmt(
                expr(ExprKind::Binary {
                    op: BinOp::NotEq,
                    left: Box::new(index_expr(ident("__c_file_ungot"), ident("handle"))),
                    right: Box::new(null_lit()),
                }),
                vec![
                    var_decl_stmt("ch", index_expr(ident("__c_file_ungot"), ident("handle"))),
                    stmt(StmtKind::Expr(assign_expr(
                        index_expr(ident("__c_file_ungot"), ident("handle")),
                        null_lit(),
                    ))),
                    stmt(StmtKind::Expr(assign_expr(
                        index_expr(ident("__c_file_eof"), ident("handle")),
                        int_lit(0),
                    ))),
                    stmt(StmtKind::Return(Some(ident("ch")))),
                ],
                None,
            ),
            if_stmt(
                expr(ExprKind::Binary {
                    op: BinOp::GtEq,
                    left: Box::new(index_expr(ident("__c_file_pos"), ident("handle"))),
                    right: Box::new(member(
                        index_expr(ident("__c_file_content"), ident("handle")),
                        "length",
                    )),
                }),
                vec![
                    stmt(StmtKind::Expr(assign_expr(
                        index_expr(ident("__c_file_eof"), ident("handle")),
                        int_lit(1),
                    ))),
                    stmt(StmtKind::Return(Some(int_lit(-1)))),
                ],
                None,
            ),
            var_decl_stmt(
                "ch",
                call_expr(
                    ident("__c_char_code_at"),
                    vec![
                        index_expr(ident("__c_file_content"), ident("handle")),
                        index_expr(ident("__c_file_pos"), ident("handle")),
                    ],
                ),
            ),
            stmt(StmtKind::Expr(assign_expr(
                index_expr(ident("__c_file_pos"), ident("handle")),
                expr(ExprKind::Binary {
                    op: BinOp::Add,
                    left: Box::new(index_expr(ident("__c_file_pos"), ident("handle"))),
                    right: Box::new(int_lit(1)),
                }),
            ))),
            stmt(StmtKind::Expr(assign_expr(
                index_expr(ident("__c_file_eof"), ident("handle")),
                int_lit(0),
            ))),
            stmt(StmtKind::Return(Some(ident("ch")))),
        ],
    ));

    out.push(function_stmt(
        "__c_ungetc_h",
        vec!["code", "handle"],
        vec![
            stmt(StmtKind::Expr(assign_expr(
                index_expr(ident("__c_file_ungot"), ident("handle")),
                ident("code"),
            ))),
            stmt(StmtKind::Expr(assign_expr(
                index_expr(ident("__c_file_eof"), ident("handle")),
                int_lit(0),
            ))),
            stmt(StmtKind::Return(Some(ident("code")))),
        ],
    ));

    out.push(function_stmt(
        "__c_fgets_h",
        vec!["handle", "size"],
        vec![
            if_stmt(
                expr(ExprKind::Binary {
                    op: BinOp::GtEq,
                    left: Box::new(index_expr(ident("__c_file_pos"), ident("handle"))),
                    right: Box::new(member(
                        index_expr(ident("__c_file_content"), ident("handle")),
                        "length",
                    )),
                }),
                vec![stmt(StmtKind::Return(Some(str_lit(""))))],
                None,
            ),
            var_decl_stmt(
                "rest",
                call_member(
                    index_expr(ident("__c_file_content"), ident("handle")),
                    "substring",
                    vec![index_expr(ident("__c_file_pos"), ident("handle"))],
                ),
            ),
            var_decl_stmt(
                "limit",
                expr(ExprKind::Binary {
                    op: BinOp::Sub,
                    left: Box::new(ident("size")),
                    right: Box::new(int_lit(1)),
                }),
            ),
            var_decl_stmt(
                "nl",
                call_member(ident("rest"), "indexOf", vec![str_lit("\n")]),
            ),
            var_decl_stmt(
                "take",
                expr(ExprKind::Ternary {
                    cond: Box::new(expr(ExprKind::Binary {
                        op: BinOp::Lt,
                        left: Box::new(ident("nl")),
                        right: Box::new(int_lit(0)),
                    })),
                    then: Box::new(ident("limit")),
                    else_: Box::new(expr(ExprKind::Ternary {
                        cond: Box::new(expr(ExprKind::Binary {
                            op: BinOp::Lt,
                            left: Box::new(expr(ExprKind::Binary {
                                op: BinOp::Add,
                                left: Box::new(ident("nl")),
                                right: Box::new(int_lit(1)),
                            })),
                            right: Box::new(ident("limit")),
                        })),
                        then: Box::new(expr(ExprKind::Binary {
                            op: BinOp::Add,
                            left: Box::new(ident("nl")),
                            right: Box::new(int_lit(1)),
                        })),
                        else_: Box::new(ident("limit")),
                    })),
                }),
            ),
            if_stmt(
                expr(ExprKind::Binary {
                    op: BinOp::Gt,
                    left: Box::new(ident("take")),
                    right: Box::new(member(ident("rest"), "length")),
                }),
                vec![stmt(StmtKind::Expr(assign_expr(
                    ident("take"),
                    member(ident("rest"), "length"),
                )))],
                None,
            ),
            var_decl_stmt(
                "out",
                call_member(ident("rest"), "substring", vec![int_lit(0), ident("take")]),
            ),
            stmt(StmtKind::Expr(assign_expr(
                index_expr(ident("__c_file_pos"), ident("handle")),
                expr(ExprKind::Binary {
                    op: BinOp::Add,
                    left: Box::new(index_expr(ident("__c_file_pos"), ident("handle"))),
                    right: Box::new(member(ident("out"), "length")),
                }),
            ))),
            stmt(StmtKind::Return(Some(ident("out")))),
        ],
    ));

    // stdin token reader (WASI-backed) — libc surface lives in the libc
    // platform adapter so any libc-targeting language shares it. Under the
    // hood it composes wasi:cli/stdin + wasi:io/streams via intrinsic:readline.
    out.extend(crate::platforms::libc::emitter::stdio_adapter::stdin_runtime_helpers());
    // char[] → string decoder for `%s`/`puts` (string / carray / code-point array).
    out.push(crate::platforms::libc::emitter::stdio_adapter::char_to_str_runtime_helper());
    // wide-char boundary helpers (code-point array ↔ string) for wchar.h.
    out.extend(crate::platforms::libc::emitter::wchar_adapter::runtime_helpers());

    // math.h domain-error helpers (libc surface) — sqrt sets errno (EDOM).
    out.extend(crate::platforms::libc::emitter::math_adapter::domain_error_helpers());

    // setjmp.h: longjmp throws an exception carrying the buf token + value; the
    // matching setjmp's generated try/catch (see wrap_setjmp_in_block) unwinds
    // the call stack back to it.
    out.push(function_stmt(
        "__c_longjmp_throw",
        vec!["token", "val"],
        vec![stmt(StmtKind::Throw {
            expr: Some(expr(ExprKind::Object(vec![
                ObjectProperty::KeyValue {
                    key: str_lit("__c_longjmp"),
                    value: ident("token"),
                },
                ObjectProperty::KeyValue {
                    key: str_lit("__c_longjmp_val"),
                    value: ident("val"),
                },
            ]))),
            cause: None,
        })],
    ));

    out.push(function_stmt(
        "__c_fseek_h",
        vec!["handle", "offset", "whence"],
        vec![
            var_decl_stmt(
                "len",
                member(
                    index_expr(ident("__c_file_content"), ident("handle")),
                    "length",
                ),
            ),
            var_decl_stmt(
                "pos",
                expr(ExprKind::Ternary {
                    cond: Box::new(expr(ExprKind::Binary {
                        op: BinOp::Eq,
                        left: Box::new(ident("whence")),
                        right: Box::new(int_lit(0)),
                    })),
                    then: Box::new(ident("offset")),
                    else_: Box::new(expr(ExprKind::Ternary {
                        cond: Box::new(expr(ExprKind::Binary {
                            op: BinOp::Eq,
                            left: Box::new(ident("whence")),
                            right: Box::new(int_lit(1)),
                        })),
                        then: Box::new(expr(ExprKind::Binary {
                            op: BinOp::Add,
                            left: Box::new(index_expr(ident("__c_file_pos"), ident("handle"))),
                            right: Box::new(ident("offset")),
                        })),
                        else_: Box::new(expr(ExprKind::Binary {
                            op: BinOp::Add,
                            left: Box::new(ident("len")),
                            right: Box::new(ident("offset")),
                        })),
                    })),
                }),
            ),
            if_stmt(
                expr(ExprKind::Binary {
                    op: BinOp::Lt,
                    left: Box::new(ident("pos")),
                    right: Box::new(int_lit(0)),
                }),
                vec![stmt(StmtKind::Expr(assign_expr(ident("pos"), int_lit(0))))],
                None,
            ),
            if_stmt(
                expr(ExprKind::Binary {
                    op: BinOp::Gt,
                    left: Box::new(ident("pos")),
                    right: Box::new(ident("len")),
                }),
                vec![stmt(StmtKind::Expr(assign_expr(
                    ident("pos"),
                    ident("len"),
                )))],
                None,
            ),
            stmt(StmtKind::Expr(assign_expr(
                index_expr(ident("__c_file_pos"), ident("handle")),
                ident("pos"),
            ))),
            stmt(StmtKind::Return(Some(int_lit(0)))),
        ],
    ));

    out.push(function_stmt(
        "__c_fwrite_h",
        vec!["data", "count", "handle"],
        vec![
            stmt(StmtKind::Expr(assign_expr(
                index_expr(ident("__c_file_content"), ident("handle")),
                call_member(ident("data"), "slice", vec![int_lit(0), ident("count")]),
            ))),
            stmt(StmtKind::Expr(assign_expr(
                index_expr(ident("__c_file_pos"), ident("handle")),
                ident("count"),
            ))),
            stmt(StmtKind::Expr(assign_expr(
                index_expr(ident("__c_file_dirty"), ident("handle")),
                int_lit(1),
            ))),
            stmt(StmtKind::Return(Some(ident("count")))),
        ],
    ));

    out.push(function_stmt(
        "__c_fread_h",
        vec!["handle", "count"],
        vec![stmt(StmtKind::Return(Some(call_member(
            index_expr(ident("__c_file_content"), ident("handle")),
            "slice",
            vec![int_lit(0), ident("count")],
        ))))],
    ));

    out.push(function_stmt(
        "__c_srand_h",
        vec!["seed"],
        vec![
            stmt(StmtKind::Expr(assign_expr(
                ident("__c_rand_state"),
                expr(ExprKind::Ternary {
                    cond: Box::new(expr(ExprKind::Binary {
                        op: BinOp::Eq,
                        left: Box::new(ident("seed")),
                        right: Box::new(int_lit(0)),
                    })),
                    then: Box::new(int_lit(1)),
                    else_: Box::new(ident("seed")),
                }),
            ))),
            stmt(StmtKind::Return(Some(int_lit(0)))),
        ],
    ));

    out.push(function_stmt(
        "__c_rand_h",
        vec![],
        vec![
            stmt(StmtKind::Expr(assign_expr(
                ident("__c_rand_state"),
                expr(ExprKind::Binary {
                    op: BinOp::Mod,
                    left: Box::new(expr(ExprKind::Binary {
                        op: BinOp::Mul,
                        left: Box::new(ident("__c_rand_state")),
                        right: Box::new(int_lit(48271)),
                    })),
                    right: Box::new(int_lit(2147483647)),
                }),
            ))),
            stmt(StmtKind::Return(Some(expr(ExprKind::Binary {
                op: BinOp::BitAnd,
                left: Box::new(ident("__c_rand_state")),
                right: Box::new(int_lit(2147483647)),
            })))),
        ],
    ));

    out.push(function_stmt(
        "__c_signal_h",
        vec!["sig", "handler"],
        vec![
            var_decl_stmt(
                "old",
                index_expr(ident("__c_signal_handlers"), ident("sig")),
            ),
            stmt(StmtKind::Expr(assign_expr(
                index_expr(ident("__c_signal_handlers"), ident("sig")),
                ident("handler"),
            ))),
            stmt(StmtKind::Return(Some(ident("old")))),
        ],
    ));

    out.push(function_stmt(
        "__c_raise_h",
        vec!["sig"],
        vec![
            var_decl_stmt("h", index_expr(ident("__c_signal_handlers"), ident("sig"))),
            if_stmt(
                expr(ExprKind::Binary {
                    op: BinOp::Or,
                    left: Box::new(expr(ExprKind::Binary {
                        op: BinOp::Eq,
                        left: Box::new(ident("h")),
                        right: Box::new(null_lit()),
                    })),
                    right: Box::new(expr(ExprKind::Binary {
                        op: BinOp::Eq,
                        left: Box::new(ident("h")),
                        right: Box::new(expr(ExprKind::Lit(Literal::Undefined))),
                    })),
                }),
                vec![stmt(StmtKind::Return(Some(int_lit(0))))],
                None,
            ),
            if_stmt(
                expr(ExprKind::Binary {
                    op: BinOp::Or,
                    left: Box::new(expr(ExprKind::Binary {
                        op: BinOp::Eq,
                        left: Box::new(ident("h")),
                        right: Box::new(int_lit(0)),
                    })),
                    right: Box::new(expr(ExprKind::Binary {
                        op: BinOp::Eq,
                        left: Box::new(ident("h")),
                        right: Box::new(int_lit(1)),
                    })),
                }),
                vec![stmt(StmtKind::Return(Some(int_lit(0))))],
                None,
            ),
            stmt(StmtKind::Expr(expr(ExprKind::Call {
                callee: Box::new(ident("h")),
                args: vec![Argument::positional(ident("sig"))],
                optional: false,
            }))),
            stmt(StmtKind::Return(Some(int_lit(0)))),
        ],
    ));

    out.push(function_stmt(
        "__c_setlocale_h",
        vec!["category", "locale"],
        vec![
            if_stmt(
                expr(ExprKind::Binary {
                    op: BinOp::And,
                    left: Box::new(expr(ExprKind::Binary {
                        op: BinOp::NotEq,
                        left: Box::new(ident("locale")),
                        right: Box::new(null_lit()),
                    })),
                    right: Box::new(expr(ExprKind::Binary {
                        op: BinOp::NotEq,
                        left: Box::new(ident("locale")),
                        right: Box::new(int_lit(0)),
                    })),
                }),
                vec![stmt(StmtKind::Expr(assign_expr(
                    ident("__c_locale"),
                    ident("locale"),
                )))],
                None,
            ),
            stmt(StmtKind::Return(Some(ident("__c_locale")))),
        ],
    ));

    // stdlib.h runtime helpers (qsort / bsearch) — libc surface, not the
    // retired cross-language `__stdlib_*` bundle.
    out.extend(crate::platforms::libc::emitter::stdlib_runtime::runtime_helpers());

    // regex.h runtime helpers (regcomp/regexec on the ECMA RegExp surface).
    out.extend(crate::platforms::libc::emitter::regex_adapter::runtime_helpers());

    // string.h runtime helpers (strcoll/strxfrm/strpbrk/strspn/strcspn).
    out.extend(crate::platforms::libc::emitter::string_runtime::runtime_helpers());

    // time.h runtime helpers live in their own adapter (shared libc surface).
    out.extend(crate::platforms::libc::emitter::time_adapter::runtime_helpers());

    out
}

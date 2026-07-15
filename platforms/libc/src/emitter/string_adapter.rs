//! C string.h — expression-level adapters for string operations.
//!
//! These helpers work with the JS-string surface (simple read-only char*).
//! The mutable carray path is handled in `pointers.rs`.

use vybe_ast::{
    Argument, BinOp, BindingPattern, ExprKind, Expression, LambdaBody, Literal, Statement,
    StmtKind, VarDeclKind, VarDeclarator,
};

use super::build::function_stmt;

fn e(kind: ExprKind) -> Expression {
    Expression::new(kind)
}

fn lit_int(n: i64) -> Expression {
    e(ExprKind::Lit(Literal::Int(n)))
}

fn member(object: Expression, field: &str) -> Expression {
    e(ExprKind::Member {
        object: Box::new(object),
        field: field.to_string(),
        null_safe: false,
    })
}

fn call(callee: Expression, args: Vec<Expression>) -> Expression {
    e(ExprKind::Call {
        callee: Box::new(callee),
        args: args.into_iter().map(Argument::positional).collect(),
        optional: false,
    })
}

fn ident(name: &str) -> Expression {
    e(ExprKind::Ident(name.to_string()))
}

fn stmt(kind: StmtKind) -> Statement {
    Statement::new(kind)
}

fn assign_expr(target: Expression, value: Expression) -> Expression {
    e(ExprKind::Assign {
        target: Box::new(target),
        value: Box::new(value),
    })
}

fn var_decl_stmt(name: &str, init: Expression) -> Statement {
    stmt(StmtKind::VarDecl {
        declarations: vec![VarDeclarator {
            pattern: BindingPattern::Ident(name.to_string()),
            type_hint: None,
            init: Some(init),
            array_bounds: None,
            with_events: false,
        }],
        kind: VarDeclKind::Var,
    })
}

fn call_member(object: Expression, field: &str, args: Vec<Expression>) -> Expression {
    call(member(object, field), args)
}

fn char_at_string(object: Expression, index: Expression) -> Expression {
    call_member(
        object,
        "substring",
        vec![
            index.clone(),
            e(ExprKind::Binary {
                op: BinOp::Add,
                left: Box::new(index),
                right: Box::new(lit_int(1)),
            }),
        ],
    )
}

fn if_stmt(
    cond: Expression,
    then_body: Vec<Statement>,
    else_body: Option<Vec<Statement>>,
) -> Statement {
    stmt(StmtKind::If {
        cond,
        then_body,
        elifs: vec![],
        else_body,
    })
}

fn lit_null() -> Expression {
    e(ExprKind::Lit(Literal::Null))
}

fn lit_str(s: &str) -> Expression {
    e(ExprKind::Lit(Literal::Str(s.to_string())))
}

/// `String.fromCharCode(code)` — int char code → 1-char string.
pub fn char_code_to_string(code: Expression) -> Expression {
    call(
        e(ExprKind::Member {
            object: Box::new(e(ExprKind::Ident("String".to_string()))),
            field: "fromCharCode".to_string(),
            null_safe: false,
        }),
        vec![code],
    )
}

/// `s.charCodeAt(0)` — first char of string → int char code.
pub fn string_to_char_code(s: Expression) -> Expression {
    call(ident("__c_char_code_at"), vec![s, lit_int(0)])
}

/// `strnlen(s, maxlen)` — bounded count until NUL or maxlen.
pub fn strnlen(s: Expression, maxlen: Expression) -> Expression {
    call(ident("__c_strnlen_h"), vec![s, maxlen])
}

/// `strchr(s, c_code)` — find first occurrence, return suffix or null.
/// `indexOf >= 0 ? s.slice(indexOf) : null`
pub fn strchr(s: Expression, c_code: Expression) -> Expression {
    let ch = char_code_to_string(c_code);
    let idx1 = call(ident("__c_str_index_of"), vec![s.clone(), ch.clone()]);
    let idx2 = call(ident("__c_str_index_of"), vec![s.clone(), ch]);
    let cond = e(ExprKind::Binary {
        op: BinOp::GtEq,
        left: Box::new(idx1),
        right: Box::new(lit_int(0)),
    });
    e(ExprKind::Ternary {
        cond: Box::new(cond),
        then: Box::new(call(member(s, "slice"), vec![idx2])),
        else_: Box::new(lit_null()),
    })
}

/// `strrchr(s, c_code)` — find last occurrence, return suffix or null.
pub fn strrchr(s: Expression, c_code: Expression) -> Expression {
    let ch = char_code_to_string(c_code);
    let idx1 = call(ident("__c_str_last_index_of"), vec![s.clone(), ch.clone()]);
    let idx2 = call(ident("__c_str_last_index_of"), vec![s.clone(), ch]);
    let cond = e(ExprKind::Binary {
        op: BinOp::GtEq,
        left: Box::new(idx1),
        right: Box::new(lit_int(0)),
    });
    e(ExprKind::Ternary {
        cond: Box::new(cond),
        then: Box::new(call(member(s, "slice"), vec![idx2])),
        else_: Box::new(lit_null()),
    })
}

/// C-facing `strchr` lowering with `needle == 0` handling.
/// For `strchr(s, '\0')` we return non-null truthy sentinel `1`.
pub fn strchr_c(s: Expression, c_code: Expression) -> Expression {
    let zero_suffix = call(
        member(s.clone(), "substring"),
        vec![member(s.clone(), "length")],
    );
    e(ExprKind::Ternary {
        cond: Box::new(e(ExprKind::Binary {
            op: BinOp::Eq,
            left: Box::new(c_code.clone()),
            right: Box::new(lit_int(0)),
        })),
        then: Box::new(zero_suffix),
        else_: Box::new(strchr(s, c_code)),
    })
}

/// C-facing `strrchr` lowering with `needle == 0` handling.
/// For `strrchr(s, '\0')` we return non-null truthy sentinel `1`.
pub fn strrchr_c(s: Expression, c_code: Expression) -> Expression {
    let zero_suffix = call(
        member(s.clone(), "substring"),
        vec![member(s.clone(), "length")],
    );
    e(ExprKind::Ternary {
        cond: Box::new(e(ExprKind::Binary {
            op: BinOp::Eq,
            left: Box::new(c_code.clone()),
            right: Box::new(lit_int(0)),
        })),
        then: Box::new(zero_suffix),
        else_: Box::new(strrchr(s, c_code)),
    })
}

/// `strstr(haystack, needle)` — find needle, return suffix or null.
pub fn strstr(haystack: Expression, needle: Expression) -> Expression {
    let idx1 = call(
        ident("__c_str_index_of"),
        vec![haystack.clone(), needle.clone()],
    );
    let idx2 = call(ident("__c_str_index_of"), vec![haystack.clone(), needle]);
    let cond = e(ExprKind::Binary {
        op: BinOp::GtEq,
        left: Box::new(idx1),
        right: Box::new(lit_int(0)),
    });
    e(ExprKind::Ternary {
        cond: Box::new(cond),
        then: Box::new(call(member(haystack, "slice"), vec![idx2])),
        else_: Box::new(lit_null()),
    })
}

/// GNU `strcasestr(haystack, needle)` — case-insensitive find, return original suffix.
pub fn strcasestr(haystack: Expression, needle: Expression) -> Expression {
    let lower_hay = call(ident("__lower__"), vec![haystack.clone()]);
    let lower_needle = call(ident("__lower__"), vec![needle]);
    let idx1 = call(
        ident("__c_str_index_of"),
        vec![lower_hay.clone(), lower_needle.clone()],
    );
    let idx2 = call(ident("__c_str_index_of"), vec![lower_hay, lower_needle]);
    let cond = e(ExprKind::Binary {
        op: BinOp::GtEq,
        left: Box::new(idx1),
        right: Box::new(lit_int(0)),
    });
    e(ExprKind::Ternary {
        cond: Box::new(cond),
        then: Box::new(call(member(haystack, "slice"), vec![idx2])),
        else_: Box::new(lit_null()),
    })
}

/// `s + n` for char pointer arithmetic on a JS string → `s.substring(n)`.
pub fn string_advance(s: Expression, n: Expression) -> Expression {
    call(member(s, "substring"), vec![n])
}

/// Stateful `strtok` lowering over two shared globals:
/// `__c_strtok_rem` and `__c_strtok_delim`.
///
/// `source_present` must encode C's `source != NULL && source != 0` check.
/// `source_value` and `delim_value` must already be normalized visible strings.
pub fn strtok_stateful(
    source_present: Expression,
    source_value: Expression,
    delim_value: Expression,
) -> Expression {
    call(
        ident("__c_strtok_h"),
        vec![source_present, source_value, delim_value],
    )
}

pub fn strtok_runtime_helpers() -> Vec<Statement> {
    vec![function_stmt(
        "__c_strtok_h",
        vec!["source_present", "source_value", "delim_value"],
        strtok_stateful_body(
            ident("source_present"),
            ident("source_value"),
            ident("delim_value"),
        ),
    )]
}

fn strtok_stateful_body(
    source_present: Expression,
    source_value: Expression,
    delim_value: Expression,
) -> Vec<Statement> {
    let leading_delim_cond = e(ExprKind::Binary {
        op: BinOp::And,
        left: Box::new(e(ExprKind::Binary {
            op: BinOp::Gt,
            left: Box::new(member(ident("rem"), "length")),
            right: Box::new(lit_int(0)),
        })),
        right: Box::new(e(ExprKind::Binary {
            op: BinOp::GtEq,
            left: Box::new(call(
                ident("__c_str_index_of"),
                vec![
                    ident("delim_text"),
                    char_at_string(ident("rem"), lit_int(0)),
                ],
            )),
            right: Box::new(lit_int(0)),
        })),
    });

    let token_scan_cond = e(ExprKind::Binary {
        op: BinOp::And,
        left: Box::new(e(ExprKind::Binary {
            op: BinOp::Lt,
            left: Box::new(ident("i")),
            right: Box::new(member(ident("rem"), "length")),
        })),
        right: Box::new(e(ExprKind::Binary {
            op: BinOp::Lt,
            left: Box::new(call(
                ident("__c_str_index_of"),
                vec![
                    ident("delim_text"),
                    char_at_string(ident("rem"), ident("i")),
                ],
            )),
            right: Box::new(lit_int(0)),
        })),
    });

    vec![
        if_stmt(
            source_present.clone(),
            vec![
                stmt(StmtKind::Expr(assign_expr(
                    ident("__c_strtok_rem"),
                    source_value,
                ))),
                stmt(StmtKind::Expr(assign_expr(
                    ident("__c_strtok_delim"),
                    delim_value,
                ))),
            ],
            None,
        ),
        var_decl_stmt("rem", ident("__c_strtok_rem")),
        var_decl_stmt("delim_text", ident("__c_strtok_delim")),
        if_stmt(
            e(ExprKind::Binary {
                op: BinOp::Eq,
                left: Box::new(ident("rem")),
                right: Box::new(lit_null()),
            }),
            vec![stmt(StmtKind::Return(Some(lit_null())))],
            None,
        ),
        if_stmt(
            e(ExprKind::Binary {
                op: BinOp::Or,
                left: Box::new(e(ExprKind::Binary {
                    op: BinOp::Eq,
                    left: Box::new(ident("delim_text")),
                    right: Box::new(lit_null()),
                })),
                right: Box::new(e(ExprKind::Binary {
                    op: BinOp::Eq,
                    left: Box::new(member(ident("delim_text"), "length")),
                    right: Box::new(lit_int(0)),
                })),
            }),
            vec![stmt(StmtKind::Expr(assign_expr(
                ident("delim_text"),
                lit_str(" "),
            )))],
            None,
        ),
        while_leading_delims(leading_delim_cond),
        if_stmt(
            e(ExprKind::Binary {
                op: BinOp::Eq,
                left: Box::new(member(ident("rem"), "length")),
                right: Box::new(lit_int(0)),
            }),
            vec![
                stmt(StmtKind::Expr(assign_expr(
                    ident("__c_strtok_rem"),
                    lit_null(),
                ))),
                stmt(StmtKind::Return(Some(lit_null()))),
            ],
            None,
        ),
        var_decl_stmt("i", lit_int(0)),
        stmt(StmtKind::While {
            cond: token_scan_cond,
            body: vec![stmt(StmtKind::Expr(assign_expr(
                ident("i"),
                e(ExprKind::Binary {
                    op: BinOp::Add,
                    left: Box::new(ident("i")),
                    right: Box::new(lit_int(1)),
                }),
            )))],
            else_body: None,
        }),
        var_decl_stmt(
            "tok",
            e(ExprKind::Ternary {
                cond: Box::new(e(ExprKind::Binary {
                    op: BinOp::Eq,
                    left: Box::new(ident("i")),
                    right: Box::new(lit_int(0)),
                })),
                then: Box::new(lit_str("")),
                else_: Box::new(call_member(
                    ident("rem"),
                    "substring",
                    vec![lit_int(0), ident("i")],
                )),
            }),
        ),
        if_stmt(
            e(ExprKind::Binary {
                op: BinOp::Lt,
                left: Box::new(ident("i")),
                right: Box::new(member(ident("rem"), "length")),
            }),
            vec![stmt(StmtKind::Expr(assign_expr(
                ident("__c_strtok_rem"),
                call_member(
                    ident("rem"),
                    "substring",
                    vec![e(ExprKind::Binary {
                        op: BinOp::Add,
                        left: Box::new(ident("i")),
                        right: Box::new(lit_int(1)),
                    })],
                ),
            )))],
            Some(vec![stmt(StmtKind::Expr(assign_expr(
                ident("__c_strtok_rem"),
                lit_null(),
            )))]),
        ),
        stmt(StmtKind::Return(Some(ident("tok")))),
    ]
}

/// Reentrant `strtok_r(source, delim, saveptr)`.
///
/// `save_target` is the C variable reached through `saveptr`; it stores the
/// remaining string between calls, so independent save variables can be nested.
pub fn strtok_r(
    source_present: Expression,
    source_value: Expression,
    delim_value: Expression,
    save_target: Expression,
) -> Expression {
    let leading_delim_cond = e(ExprKind::Binary {
        op: BinOp::And,
        left: Box::new(e(ExprKind::Binary {
            op: BinOp::Gt,
            left: Box::new(member(ident("rem"), "length")),
            right: Box::new(lit_int(0)),
        })),
        right: Box::new(e(ExprKind::Binary {
            op: BinOp::GtEq,
            left: Box::new(call(
                ident("__c_str_index_of"),
                vec![
                    ident("delim_text"),
                    char_at_string(ident("rem"), lit_int(0)),
                ],
            )),
            right: Box::new(lit_int(0)),
        })),
    });

    let token_scan_cond = e(ExprKind::Binary {
        op: BinOp::And,
        left: Box::new(e(ExprKind::Binary {
            op: BinOp::Lt,
            left: Box::new(ident("i")),
            right: Box::new(member(ident("rem"), "length")),
        })),
        right: Box::new(e(ExprKind::Binary {
            op: BinOp::Lt,
            left: Box::new(call(
                ident("__c_str_index_of"),
                vec![
                    ident("delim_text"),
                    char_at_string(ident("rem"), ident("i")),
                ],
            )),
            right: Box::new(lit_int(0)),
        })),
    });

    e(ExprKind::Call {
        callee: Box::new(e(ExprKind::Lambda {
            params: vec![],
            body: LambdaBody::Block(vec![
                if_stmt(
                    source_present.clone(),
                    vec![stmt(StmtKind::Expr(assign_expr(
                        save_target.clone(),
                        source_value,
                    )))],
                    None,
                ),
                var_decl_stmt("rem", save_target.clone()),
                var_decl_stmt("delim_text", delim_value),
                if_stmt(
                    e(ExprKind::Binary {
                        op: BinOp::Eq,
                        left: Box::new(ident("rem")),
                        right: Box::new(lit_null()),
                    }),
                    vec![stmt(StmtKind::Return(Some(lit_null())))],
                    None,
                ),
                if_stmt(
                    e(ExprKind::Binary {
                        op: BinOp::Or,
                        left: Box::new(e(ExprKind::Binary {
                            op: BinOp::Eq,
                            left: Box::new(ident("delim_text")),
                            right: Box::new(lit_null()),
                        })),
                        right: Box::new(e(ExprKind::Binary {
                            op: BinOp::Eq,
                            left: Box::new(member(ident("delim_text"), "length")),
                            right: Box::new(lit_int(0)),
                        })),
                    }),
                    vec![stmt(StmtKind::Expr(assign_expr(
                        ident("delim_text"),
                        lit_str(" "),
                    )))],
                    None,
                ),
                while_leading_delims(leading_delim_cond),
                if_stmt(
                    e(ExprKind::Binary {
                        op: BinOp::Eq,
                        left: Box::new(member(ident("rem"), "length")),
                        right: Box::new(lit_int(0)),
                    }),
                    vec![
                        stmt(StmtKind::Expr(assign_expr(save_target.clone(), lit_null()))),
                        stmt(StmtKind::Return(Some(lit_null()))),
                    ],
                    None,
                ),
                var_decl_stmt("i", lit_int(0)),
                stmt(StmtKind::While {
                    cond: token_scan_cond,
                    body: vec![stmt(StmtKind::Expr(assign_expr(
                        ident("i"),
                        e(ExprKind::Binary {
                            op: BinOp::Add,
                            left: Box::new(ident("i")),
                            right: Box::new(lit_int(1)),
                        }),
                    )))],
                    else_body: None,
                }),
                var_decl_stmt(
                    "tok",
                    e(ExprKind::Ternary {
                        cond: Box::new(e(ExprKind::Binary {
                            op: BinOp::Eq,
                            left: Box::new(ident("i")),
                            right: Box::new(lit_int(0)),
                        })),
                        then: Box::new(lit_str("")),
                        else_: Box::new(call_member(
                            ident("rem"),
                            "substring",
                            vec![lit_int(0), ident("i")],
                        )),
                    }),
                ),
                if_stmt(
                    e(ExprKind::Binary {
                        op: BinOp::Lt,
                        left: Box::new(ident("i")),
                        right: Box::new(member(ident("rem"), "length")),
                    }),
                    vec![stmt(StmtKind::Expr(assign_expr(
                        save_target.clone(),
                        call_member(
                            ident("rem"),
                            "substring",
                            vec![e(ExprKind::Binary {
                                op: BinOp::Add,
                                left: Box::new(ident("i")),
                                right: Box::new(lit_int(1)),
                            })],
                        ),
                    )))],
                    Some(vec![stmt(StmtKind::Expr(assign_expr(
                        save_target,
                        lit_null(),
                    )))]),
                ),
                stmt(StmtKind::Return(Some(ident("tok")))),
            ]),
            is_async: false,
            captures: vec![],
        })),
        args: vec![],
        optional: false,
    })
}

fn while_leading_delims(cond: Expression) -> Statement {
    stmt(StmtKind::While {
        cond,
        body: vec![stmt(StmtKind::Expr(assign_expr(
            ident("rem"),
            call_member(ident("rem"), "substring", vec![lit_int(1)]),
        )))],
        else_body: None,
    })
}

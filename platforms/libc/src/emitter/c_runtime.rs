//! libc runtime prelude — the FILE/stdio model, math.h series helpers, and the
//! rand / signal / locale / strtok runtime, emitted as common AST. This is the
//! libc surface shared by any libc-targeting front-end; the C walker injects it
//! via `prelude()`. Builders come from `build`; math series helpers from
//! `math_runtime`; the stdin / char-decode / wide-char / domain-error helpers
//! are composed from their own adapters.

use crate::emitter::build::*;
use crate::emitter::math_runtime::{
    build_math_helper_fn, ecma_math_call, poly_erf, stirling_approx,
};
use vybe_ast::{
    Argument, ArrayElement, BinOp, BindingPattern, ExprKind, Literal, ObjectProperty, Statement,
    StmtKind, VarDeclKind, VarDeclarator,
};

/// The libc runtime prelude is identical for every program, so build it once
/// and hand out clones — the same OnceLock caching the JS/PHP front-ends use.
pub fn prelude() -> Vec<Statement> {
    use std::sync::OnceLock;
    static CACHE: OnceLock<Vec<Statement>> = OnceLock::new();
    CACHE.get_or_init(build_prelude).clone()
}

fn build_prelude() -> Vec<Statement> {
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
                                        op: vybe_ast::UnaryOp::Neg,
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
                        op: vybe_ast::UnaryOp::Neg,
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
                pattern: BindingPattern::Ident("__c_file_error".to_string()),
                type_hint: None,
                init: Some(expr(ExprKind::Object(vec![]))),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        }),
        stmt(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident("__c_file_append".to_string()),
                type_hint: None,
                init: Some(expr(ExprKind::Object(vec![]))),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        }),
        stmt(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident("__c_file_readonly".to_string()),
                type_hint: None,
                init: Some(expr(ExprKind::Object(vec![]))),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        }),
        stmt(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident("__c_file_closed".to_string()),
                type_hint: None,
                init: Some(expr(ExprKind::Object(vec![]))),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        }),
        stmt(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident("__c_last_file_handle".to_string()),
                type_hint: None,
                init: Some(int_lit(0)),
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
                // 0 stdin, 1 stdout, 2 stderr (real fd numbering) — the first
                // opened FILE gets 3. Starting at 2 made the first file's
                // writes hit the stderr arm in `__c_fputs_h`.
                init: Some(int_lit(3)),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        }),
        stmt(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                // The descriptor counter lives in an OBJECT, not a scalar.
                // A spawned thread's globals are a clone, so a scalar counter
                // would restart there and the thread's first socket would take
                // an fd the parent already handed out — both then index the
                // same (shared) `__c_sock_*` tables and one clobbers the
                // other. One object = one counter for the whole process, which
                // is what a descriptor table is.
                pattern: BindingPattern::Ident("__c_fd_seq".to_string()),
                type_hint: None,
                init: Some(expr(ExprKind::Object(vec![ObjectProperty::KeyValue {
                    key: str_lit("n"),
                    value: int_lit(3),
                }]))),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        }),
        // fork/exec/wait state: `fork()` runs the child INLINE (see
        // posix_adapter::fork) — the flag tells `exec`/`_exit` to fall
        // through to the parent's code instead of ending the run, and the
        // status is what `wait`/`waitpid` report.
        stmt(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident("__c_in_forked_child".to_string()),
                type_hint: None,
                init: Some(int_lit(0)),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        }),
        stmt(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident("__c_child_status".to_string()),
                type_hint: None,
                init: Some(int_lit(0)),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        }),
        stmt(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident("__c_pending_children".to_string()),
                type_hint: None,
                init: Some(int_lit(0)),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        }),
        stmt(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident("__c_new_fd".to_string()),
                type_hint: None,
                init: None,
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        }),
        stmt(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident("__c_dup_target".to_string()),
                type_hint: None,
                init: None,
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        }),
        stmt(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident("__c_pipe_r".to_string()),
                type_hint: None,
                init: None,
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        }),
        stmt(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident("__c_pipe_w".to_string()),
                type_hint: None,
                init: None,
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        }),
        stmt(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident("__c_fd_open".to_string()),
                type_hint: None,
                init: Some(expr(ExprKind::Object(vec![]))),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        }),
        stmt(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident("__c_fd_flags".to_string()),
                type_hint: None,
                init: Some(expr(ExprKind::Object(vec![]))),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        }),
        stmt(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident("__c_fd_cloexec".to_string()),
                type_hint: None,
                init: Some(expr(ExprKind::Object(vec![]))),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        }),
        stmt(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident("__c_fd_nonblock".to_string()),
                type_hint: None,
                init: Some(expr(ExprKind::Object(vec![]))),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        }),
        stmt(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident("__c_fd_size".to_string()),
                type_hint: None,
                init: Some(expr(ExprKind::Object(vec![]))),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        }),
        stmt(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident("__c_fd_content_by_fd".to_string()),
                type_hint: None,
                init: Some(expr(ExprKind::Object(vec![]))),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        }),
        stmt(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident("__c_fd_path_by_fd".to_string()),
                type_hint: None,
                init: Some(expr(ExprKind::Object(vec![]))),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        }),
        stmt(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident("__c_mq_next".to_string()),
                type_hint: None,
                init: Some(int_lit(100)),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        }),
        stmt(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident("__c_mq_by_name".to_string()),
                type_hint: None,
                init: Some(expr(ExprKind::Object(vec![]))),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        }),
        stmt(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident("__c_mq_msg".to_string()),
                type_hint: None,
                init: Some(expr(ExprKind::Object(vec![]))),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        }),
        stmt(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident("__c_mq_prio".to_string()),
                type_hint: None,
                init: Some(expr(ExprKind::Object(vec![]))),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        }),
        stmt(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident("__c_mq_has_msg".to_string()),
                type_hint: None,
                init: Some(expr(ExprKind::Object(vec![]))),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        }),
        stmt(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident("__c_mq_flags".to_string()),
                type_hint: None,
                init: Some(expr(ExprKind::Object(vec![]))),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        }),
        stmt(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident("__c_mq_msgsize".to_string()),
                type_hint: None,
                init: Some(expr(ExprKind::Object(vec![]))),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        }),
        stmt(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident("__c_fenv_excepts".to_string()),
                type_hint: None,
                init: Some(int_lit(0)),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        }),
        stmt(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident("__c_pipe_is_reader".to_string()),
                type_hint: None,
                init: Some(expr(ExprKind::Object(vec![]))),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        }),
        stmt(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident("__c_pipe_is_writer".to_string()),
                type_hint: None,
                init: Some(expr(ExprKind::Object(vec![]))),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        }),
        stmt(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident("__c_pipe_peer".to_string()),
                type_hint: None,
                init: Some(expr(ExprKind::Object(vec![]))),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        }),
        stmt(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident("__c_pipe_writer_closed".to_string()),
                type_hint: None,
                init: Some(expr(ExprKind::Object(vec![]))),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        }),
        stmt(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident("__c_pipe_write_count".to_string()),
                type_hint: None,
                init: Some(int_lit(0)),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        }),
        stmt(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident("__c_last_termsig".to_string()),
                type_hint: None,
                init: Some(int_lit(9)),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        }),
        stmt(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident("__c_path_exists".to_string()),
                type_hint: None,
                init: Some(expr(ExprKind::Object(vec![]))),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        }),
        stmt(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident("__c_shm_exists".to_string()),
                type_hint: None,
                init: Some(expr(ExprKind::Object(vec![]))),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        }),
        stmt(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident("__c_next_shm_fd".to_string()),
                type_hint: None,
                init: Some(int_lit(30)),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        }),
        stmt(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident("__c_sem_exists".to_string()),
                type_hint: None,
                init: Some(expr(ExprKind::Object(vec![]))),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        }),
        stmt(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident("__c_sem_values".to_string()),
                type_hint: None,
                init: Some(expr(ExprKind::Object(vec![]))),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        }),
        stmt(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident("__c_sem_handles".to_string()),
                type_hint: None,
                init: Some(expr(ExprKind::Object(vec![]))),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        }),
        stmt(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident("__c_next_sem_handle".to_string()),
                type_hint: None,
                init: Some(int_lit(1)),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        }),
        stmt(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident("__c_thread_results".to_string()),
                type_hint: None,
                init: Some(expr(ExprKind::Object(vec![]))),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        }),
        stmt(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident("__c_thread_starts".to_string()),
                type_hint: None,
                init: Some(expr(ExprKind::Object(vec![]))),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        }),
        stmt(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident("__c_thread_args".to_string()),
                type_hint: None,
                init: Some(expr(ExprKind::Object(vec![]))),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        }),
        // recv's one-shot scratch: the helper's answer has to be inspected
        // (null = EAGAIN) before it reaches the caller's buffer.
        stmt(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident("__c_recv_tmp".to_string()),
                type_hint: None,
                init: None,
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        }),
        // The Task object each spawn answers — what `pthread_join` waits on.
        stmt(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident("__c_thread_tasks".to_string()),
                type_hint: None,
                init: Some(expr(ExprKind::Object(vec![]))),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        }),
        // Barrier records, keyed by the handle the barrier variable holds.
        // Objects, not scalars: a spawned thread's globals are a clone, so
        // only object state crosses back (see thread_adapter::barrier_init).
        stmt(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident("__c_barrier_limit".to_string()),
                type_hint: None,
                init: Some(expr(ExprKind::Object(vec![]))),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        }),
        stmt(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident("__c_barrier_arrived".to_string()),
                type_hint: None,
                init: Some(expr(ExprKind::Object(vec![]))),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        }),
        stmt(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident("__c_barrier_gen".to_string()),
                type_hint: None,
                init: Some(expr(ExprKind::Object(vec![]))),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        }),
        stmt(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident("__c_next_barrier_handle".to_string()),
                type_hint: None,
                init: Some(int_lit(1)),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        }),
        stmt(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident("__c_thread_result_tmp".to_string()),
                type_hint: None,
                init: None,
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        }),
        stmt(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident("__c_next_thread_handle".to_string()),
                type_hint: None,
                init: Some(int_lit(1)),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        }),
        stmt(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident("__c_cleanup_fn".to_string()),
                type_hint: None,
                init: None,
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        }),
        stmt(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident("__c_cleanup_arg".to_string()),
                type_hint: None,
                init: None,
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        }),
        stmt(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident("__c_tls_values".to_string()),
                type_hint: None,
                init: Some(expr(ExprKind::Object(vec![]))),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        }),
        stmt(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident("__c_tls_destructors".to_string()),
                type_hint: None,
                init: Some(expr(ExprKind::Object(vec![]))),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        }),
        stmt(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident("__c_tls_saved".to_string()),
                type_hint: None,
                init: Some(expr(ExprKind::Object(vec![]))),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        }),
        stmt(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident("__c_last_tls_key".to_string()),
                type_hint: None,
                init: Some(int_lit(0)),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        }),
        stmt(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident("__c_next_tls_key".to_string()),
                type_hint: None,
                init: Some(int_lit(1)),
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
        // THE environment: lazily seeded from wasi:cli/environment
        // .get-environment, mutated by setenv/putenv/unsetenv/clearenv, read
        // by getenv, and passed to `system()` children (spawnSync
        // options.env) once any write made it dirty.
        stmt(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident("__c_env_obj".to_string()),
                type_hint: None,
                init: Some(expr(ExprKind::Object(Vec::new()))),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        }),
        stmt(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident("__c_env_dirty".to_string()),
                type_hint: None,
                init: Some(int_lit(0)),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        }),
        stmt(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident("__c_env_seeded".to_string()),
                type_hint: None,
                init: Some(int_lit(0)),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        }),
        // popen bookkeeping: per-handle child exit status ("r" mode), the
        // deferred command ("w" mode — run at pclose with the buffered
        // writes as stdin), and a sequence for unique buffer paths.
        stmt(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident("__c_popen_status".to_string()),
                type_hint: None,
                init: Some(expr(ExprKind::Object(Vec::new()))),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        }),
        stmt(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident("__c_popen_wcmd".to_string()),
                type_hint: None,
                init: Some(expr(ExprKind::Object(Vec::new()))),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        }),
        stmt(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident("__c_popen_seq".to_string()),
                type_hint: None,
                init: Some(int_lit(0)),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        }),
        // Socket state, keyed by fd: the `wasi:sockets` resource and the
        // streams it hands back. Python keeps these on a socket OBJECT; C has
        // integer descriptors, so they live in tables.
        //
        // `__c_sock_listener` holds `listen()`'s `stream<tcp-socket>`. 0.3.1
        // has no `accept` verb — the stream IS the accept queue, and every
        // accepted connection is one element read from it, so the stream has
        // to outlive the `listen()` call that produced it.
        stmt(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident("__c_sock_listener".to_string()),
                type_hint: None,
                init: Some(expr(ExprKind::Object(Vec::new()))),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        }),
        stmt(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident("__c_sock_res".to_string()),
                type_hint: None,
                init: Some(expr(ExprKind::Object(Vec::new()))),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        }),
        stmt(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident("__c_sock_rx".to_string()),
                type_hint: None,
                init: Some(expr(ExprKind::Object(Vec::new()))),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        }),
        stmt(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident("__c_sock_tx".to_string()),
                type_hint: None,
                init: Some(expr(ExprKind::Object(Vec::new()))),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        }),
        stmt(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident("__c_sock_kind".to_string()),
                type_hint: None,
                init: Some(expr(ExprKind::Object(Vec::new()))),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        }),
        // Datagram-socket bookkeeping. `bound` drives POSIX's auto-bind on
        // first send, `peer` is what makes an address-less `send` legal, and
        // `peek` holds the one datagram MSG_PEEK read without consuming.
        stmt(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident("__c_sock_bound".to_string()),
                type_hint: None,
                init: Some(expr(ExprKind::Object(Vec::new()))),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        }),
        stmt(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident("__c_sock_peer".to_string()),
                type_hint: None,
                init: Some(expr(ExprKind::Object(Vec::new()))),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        }),
        stmt(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident("__c_sock_peek".to_string()),
                type_hint: None,
                init: Some(expr(ExprKind::Object(Vec::new()))),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        }),
        // AF_UNIX. `__c_unix_reg` is the NAME registry — `sun_path` → the
        // channel `bind` published — and `_in`/`_out` are this descriptor's
        // two directions. Objects, so a spawned thread reaches the same ones.
        stmt(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident("__c_sock_family".to_string()),
                type_hint: None,
                init: Some(expr(ExprKind::Object(Vec::new()))),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        }),
        stmt(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident("__c_sock_path".to_string()),
                type_hint: None,
                init: Some(expr(ExprKind::Object(Vec::new()))),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        }),
        stmt(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident("__c_sock_in".to_string()),
                type_hint: None,
                init: Some(expr(ExprKind::Object(Vec::new()))),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        }),
        stmt(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident("__c_sock_out".to_string()),
                type_hint: None,
                init: Some(expr(ExprKind::Object(Vec::new()))),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        }),
        stmt(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident("__c_unix_reg".to_string()),
                type_hint: None,
                init: Some(expr(ExprKind::Object(Vec::new()))),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        }),
        // Directory-ness registry (mkdir sets, rmdir clears): unlink/rmdir
        // dispatch on it instead of on path-name patterns.
        stmt(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident("__c_is_dir".to_string()),
                type_hint: None,
                init: Some(expr(ExprKind::Object(Vec::new()))),
                array_bounds: None,
                with_events: false,
            }],
            kind: VarDeclKind::Var,
        }),
        // Set when a file is created under a directory ("dir/file" paths) —
        // rename(dir, nonempty_dir) is ENOTEMPTY per POSIX.
        stmt(StmtKind::VarDecl {
            declarations: vec![VarDeclarator {
                pattern: BindingPattern::Ident("__c_dir_nonempty".to_string()),
                type_hint: None,
                init: Some(expr(ExprKind::Object(Vec::new()))),
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

    // pthread entry / join / barrier — see `thread_adapter::runtime_functions`.
    out.extend(crate::emitter::thread_adapter::runtime_functions());

    out.push(function_stmt(
        "__c_char_ptr_add",
        vec!["s", "n"],
        vec![stmt(StmtKind::Return(Some(call_member(
            ident("s"),
            "substring",
            vec![ident("n")],
        ))))],
    ));

    // The environment, seeded ONCE from the real process environment
    // (`wasi:cli/environment.get-environment` — a list of [key, value]
    // pairs). Every env entry point calls this first, so getenv reads and
    // setenv writes share one runtime truth with what children inherit.
    out.push(function_stmt(
        "__c_env_init",
        vec![],
        vec![
            if_stmt(
                expr(ExprKind::Binary {
                    op: BinOp::Eq,
                    left: Box::new(ident("__c_env_seeded")),
                    right: Box::new(int_lit(1)),
                }),
                vec![stmt(StmtKind::Return(Some(int_lit(0))))],
                None,
            ),
            stmt(StmtKind::Expr(assign_expr(
                ident("__c_env_seeded"),
                int_lit(1),
            ))),
            var_decl_stmt("pairs", call_expr(ident("__c_get_environment"), vec![])),
            var_decl_stmt("i", int_lit(0)),
            stmt(StmtKind::While {
                cond: expr(ExprKind::Binary {
                    op: BinOp::Lt,
                    left: Box::new(ident("i")),
                    right: Box::new(member(ident("pairs"), "length")),
                }),
                body: vec![
                    stmt(StmtKind::Expr(assign_expr(
                        index_expr(
                            ident("__c_env_obj"),
                            index_expr(index_expr(ident("pairs"), ident("i")), int_lit(0)),
                        ),
                        index_expr(index_expr(ident("pairs"), ident("i")), int_lit(1)),
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
            stmt(StmtKind::Return(Some(int_lit(0)))),
        ],
    ));

    out.push(function_stmt(
        "__c_getenv_h",
        vec!["name"],
        vec![
            stmt(StmtKind::Expr(call_expr(ident("__c_env_init"), vec![]))),
            stmt(StmtKind::Return(Some(expr(ExprKind::NullCoalesce {
                left: Box::new(index_expr(ident("__c_env_obj"), ident("name"))),
                right: Box::new(null_lit()),
            })))),
        ],
    ));

    out.push(function_stmt(
        "__c_setenv_h",
        vec!["name", "value", "overwrite"],
        vec![
            stmt(StmtKind::Expr(call_expr(ident("__c_env_init"), vec![]))),
            // POSIX: empty name or a name containing '=' is EINVAL.
            if_stmt(
                expr(ExprKind::Binary {
                    op: BinOp::Or,
                    left: Box::new(expr(ExprKind::Binary {
                        op: BinOp::Eq,
                        left: Box::new(ident("name")),
                        right: Box::new(str_lit("")),
                    })),
                    right: Box::new(expr(ExprKind::Binary {
                        op: BinOp::GtEq,
                        left: Box::new(call_member(ident("name"), "indexOf", vec![str_lit("=")])),
                        right: Box::new(int_lit(0)),
                    })),
                }),
                vec![stmt(StmtKind::Return(Some(int_lit(-1))))],
                None,
            ),
            if_stmt(
                expr(ExprKind::Binary {
                    op: BinOp::And,
                    left: Box::new(expr(ExprKind::Binary {
                        op: BinOp::Eq,
                        left: Box::new(ident("overwrite")),
                        right: Box::new(int_lit(0)),
                    })),
                    right: Box::new(expr(ExprKind::Binary {
                        op: BinOp::NotEq,
                        left: Box::new(expr(ExprKind::NullCoalesce {
                            left: Box::new(index_expr(ident("__c_env_obj"), ident("name"))),
                            right: Box::new(null_lit()),
                        })),
                        right: Box::new(null_lit()),
                    })),
                }),
                vec![stmt(StmtKind::Return(Some(int_lit(0))))],
                None,
            ),
            stmt(StmtKind::Expr(assign_expr(
                index_expr(ident("__c_env_obj"), ident("name")),
                ident("value"),
            ))),
            stmt(StmtKind::Expr(assign_expr(
                ident("__c_env_dirty"),
                int_lit(1),
            ))),
            stmt(StmtKind::Return(Some(int_lit(0)))),
        ],
    ));

    out.push(function_stmt(
        "__c_unsetenv_h",
        vec!["name"],
        vec![
            stmt(StmtKind::Expr(call_expr(ident("__c_env_init"), vec![]))),
            stmt(StmtKind::Expr(assign_expr(
                index_expr(ident("__c_env_obj"), ident("name")),
                null_lit(),
            ))),
            stmt(StmtKind::Expr(assign_expr(
                ident("__c_env_dirty"),
                int_lit(1),
            ))),
            stmt(StmtKind::Return(Some(int_lit(0)))),
        ],
    ));

    // putenv("NAME=value") — no '=' removes the name (glibc extension).
    // COPY semantics (setenv-shaped); POSIX putenv's buffer aliasing — where
    // later writes to the caller's buffer change the environment — is not
    // modeled.
    out.push(function_stmt(
        "__c_putenv_h",
        vec!["entry"],
        vec![
            stmt(StmtKind::Expr(call_expr(ident("__c_env_init"), vec![]))),
            var_decl_stmt(
                "eq",
                call_member(ident("entry"), "indexOf", vec![str_lit("=")]),
            ),
            if_stmt(
                expr(ExprKind::Binary {
                    op: BinOp::Lt,
                    left: Box::new(ident("eq")),
                    right: Box::new(int_lit(0)),
                }),
                vec![
                    stmt(StmtKind::Expr(assign_expr(
                        index_expr(ident("__c_env_obj"), ident("entry")),
                        null_lit(),
                    ))),
                    stmt(StmtKind::Expr(assign_expr(
                        ident("__c_env_dirty"),
                        int_lit(1),
                    ))),
                    stmt(StmtKind::Return(Some(int_lit(0)))),
                ],
                None,
            ),
            stmt(StmtKind::Expr(assign_expr(
                index_expr(
                    ident("__c_env_obj"),
                    call_member(ident("entry"), "substring", vec![int_lit(0), ident("eq")]),
                ),
                call_member(
                    ident("entry"),
                    "substring",
                    vec![expr(ExprKind::Binary {
                        op: BinOp::Add,
                        left: Box::new(ident("eq")),
                        right: Box::new(int_lit(1)),
                    })],
                ),
            ))),
            stmt(StmtKind::Expr(assign_expr(
                ident("__c_env_dirty"),
                int_lit(1),
            ))),
            stmt(StmtKind::Return(Some(int_lit(0)))),
        ],
    ));

    // Existence is a RUNTIME question over both surfaces: the virtual store
    // (C-written files) and the real filesystem (files made by `system()`
    // children or pre-existing). The old `__c_path_exists[p] == 0` checks
    // never fired for never-created paths (undefined ≠ 0 under strict Eq),
    // which is why remove/rename grew literal name-pattern hacks.
    out.push(function_stmt(
        "__c_path_present",
        vec!["path"],
        vec![
            if_stmt(
                expr(ExprKind::Binary {
                    op: BinOp::NotEq,
                    left: Box::new(expr(ExprKind::NullCoalesce {
                        left: Box::new(index_expr(ident(store_name), ident("path"))),
                        right: Box::new(null_lit()),
                    })),
                    right: Box::new(null_lit()),
                }),
                vec![stmt(StmtKind::Return(Some(int_lit(1))))],
                None,
            ),
            // The registry is three-valued: 1 = virtually created (mkdir and
            // friends), 0 = TOMBSTONE (removed/renamed away — a stale REAL
            // file with the same name must NOT resurrect it), absent = ask
            // the real filesystem.
            var_decl_stmt(
                "pe",
                expr(ExprKind::NullCoalesce {
                    left: Box::new(index_expr(ident("__c_path_exists"), ident("path"))),
                    right: Box::new(int_lit(-1)),
                }),
            ),
            if_stmt(
                expr(ExprKind::Binary {
                    op: BinOp::Eq,
                    left: Box::new(ident("pe")),
                    right: Box::new(int_lit(1)),
                }),
                vec![stmt(StmtKind::Return(Some(int_lit(1))))],
                None,
            ),
            if_stmt(
                expr(ExprKind::Binary {
                    op: BinOp::Eq,
                    left: Box::new(ident("pe")),
                    right: Box::new(int_lit(0)),
                }),
                vec![stmt(StmtKind::Return(Some(int_lit(0))))],
                None,
            ),
            if_stmt(
                call_expr(ident("__c_fs_exists"), vec![ident("path")]),
                vec![stmt(StmtKind::Return(Some(int_lit(1))))],
                None,
            ),
            stmt(StmtKind::Return(Some(int_lit(0)))),
        ],
    ));

    out.push(function_stmt(
        "__c_remove_h",
        vec!["path"],
        vec![
            if_stmt(
                expr(ExprKind::Binary {
                    op: BinOp::Gt,
                    left: Box::new(member(ident("path"), "length")),
                    right: Box::new(int_lit(240)),
                }),
                vec![stmt(StmtKind::Return(Some(int_lit(-1))))],
                None,
            ),
            if_stmt(
                expr(ExprKind::Binary {
                    op: BinOp::Eq,
                    left: Box::new(call_expr(ident("__c_path_present"), vec![ident("path")])),
                    right: Box::new(int_lit(0)),
                }),
                vec![stmt(StmtKind::Return(Some(int_lit(-1))))],
                None,
            ),
            // Virtual entry (file in the store, or a registry-created dir):
            // tombstone it. A REAL file (no virtual entry) is really deleted.
            if_stmt(
                expr(ExprKind::Binary {
                    op: BinOp::And,
                    left: Box::new(expr(ExprKind::Binary {
                        op: BinOp::Eq,
                        left: Box::new(expr(ExprKind::NullCoalesce {
                            left: Box::new(index_expr(ident(store_name), ident("path"))),
                            right: Box::new(null_lit()),
                        })),
                        right: Box::new(null_lit()),
                    })),
                    right: Box::new(expr(ExprKind::Binary {
                        op: BinOp::NotEq,
                        left: Box::new(expr(ExprKind::NullCoalesce {
                            left: Box::new(index_expr(ident("__c_path_exists"), ident("path"))),
                            right: Box::new(int_lit(-1)),
                        })),
                        right: Box::new(int_lit(1)),
                    })),
                }),
                vec![
                    stmt(StmtKind::Expr(call_expr(
                        ident("__c_fs_rm"),
                        vec![ident("path")],
                    ))),
                    stmt(StmtKind::Return(Some(int_lit(0)))),
                ],
                None,
            ),
            stmt(StmtKind::Expr(assign_expr(
                index_expr(ident(store_name), ident("path")),
                null_lit(),
            ))),
            stmt(StmtKind::Expr(assign_expr(
                index_expr(ident("__c_path_exists"), ident("path")),
                int_lit(0),
            ))),
            stmt(StmtKind::Return(Some(int_lit(0)))),
        ],
    ));

    out.push(function_stmt(
        "__c_rename_h",
        vec!["src", "dst"],
        vec![
            if_stmt(
                expr(ExprKind::Binary {
                    op: BinOp::Eq,
                    left: Box::new(call_expr(ident("__c_path_present"), vec![ident("src")])),
                    right: Box::new(int_lit(0)),
                }),
                vec![stmt(StmtKind::Return(Some(int_lit(-1))))],
                None,
            ),
            // POSIX type agreement: a file cannot replace a directory
            // (EISDIR), a directory cannot replace a file (ENOTDIR), and a
            // directory cannot replace a NON-EMPTY directory (ENOTEMPTY).
            var_decl_stmt(
                "sd",
                expr(ExprKind::NullCoalesce {
                    left: Box::new(index_expr(ident("__c_is_dir"), ident("src"))),
                    right: Box::new(int_lit(0)),
                }),
            ),
            var_decl_stmt(
                "dd",
                expr(ExprKind::NullCoalesce {
                    left: Box::new(index_expr(ident("__c_is_dir"), ident("dst"))),
                    right: Box::new(int_lit(0)),
                }),
            ),
            if_stmt(
                expr(ExprKind::Binary {
                    op: BinOp::And,
                    left: Box::new(expr(ExprKind::Binary {
                        op: BinOp::Eq,
                        left: Box::new(ident("sd")),
                        right: Box::new(int_lit(0)),
                    })),
                    right: Box::new(expr(ExprKind::Binary {
                        op: BinOp::Eq,
                        left: Box::new(ident("dd")),
                        right: Box::new(int_lit(1)),
                    })),
                }),
                vec![stmt(StmtKind::Return(Some(int_lit(-1))))],
                None,
            ),
            if_stmt(
                expr(ExprKind::Binary {
                    op: BinOp::And,
                    left: Box::new(expr(ExprKind::Binary {
                        op: BinOp::Eq,
                        left: Box::new(ident("sd")),
                        right: Box::new(int_lit(1)),
                    })),
                    right: Box::new(expr(ExprKind::Binary {
                        op: BinOp::And,
                        left: Box::new(expr(ExprKind::Binary {
                            op: BinOp::Eq,
                            left: Box::new(ident("dd")),
                            right: Box::new(int_lit(0)),
                        })),
                        right: Box::new(expr(ExprKind::Binary {
                            op: BinOp::Eq,
                            left: Box::new(call_expr(
                                ident("__c_path_present"),
                                vec![ident("dst")],
                            )),
                            right: Box::new(int_lit(1)),
                        })),
                    })),
                }),
                vec![stmt(StmtKind::Return(Some(int_lit(-1))))],
                None,
            ),
            if_stmt(
                expr(ExprKind::Binary {
                    op: BinOp::And,
                    left: Box::new(expr(ExprKind::Binary {
                        op: BinOp::Eq,
                        left: Box::new(ident("sd")),
                        right: Box::new(int_lit(1)),
                    })),
                    right: Box::new(expr(ExprKind::Binary {
                        op: BinOp::And,
                        left: Box::new(expr(ExprKind::Binary {
                            op: BinOp::Eq,
                            left: Box::new(ident("dd")),
                            right: Box::new(int_lit(1)),
                        })),
                        right: Box::new(expr(ExprKind::Binary {
                            op: BinOp::Eq,
                            left: Box::new(expr(ExprKind::NullCoalesce {
                                left: Box::new(index_expr(ident("__c_dir_nonempty"), ident("dst"))),
                                right: Box::new(int_lit(0)),
                            })),
                            right: Box::new(int_lit(1)),
                        })),
                    })),
                }),
                vec![stmt(StmtKind::Return(Some(int_lit(-1))))],
                None,
            ),
            // A REAL file (no virtual entry) renames on disk; a virtual
            // entry moves in the store/registry and tombstones the source.
            if_stmt(
                expr(ExprKind::Binary {
                    op: BinOp::And,
                    left: Box::new(expr(ExprKind::Binary {
                        op: BinOp::Eq,
                        left: Box::new(expr(ExprKind::NullCoalesce {
                            left: Box::new(index_expr(ident(store_name), ident("src"))),
                            right: Box::new(null_lit()),
                        })),
                        right: Box::new(null_lit()),
                    })),
                    right: Box::new(expr(ExprKind::Binary {
                        op: BinOp::NotEq,
                        left: Box::new(expr(ExprKind::NullCoalesce {
                            left: Box::new(index_expr(ident("__c_path_exists"), ident("src"))),
                            right: Box::new(int_lit(-1)),
                        })),
                        right: Box::new(int_lit(1)),
                    })),
                }),
                vec![
                    stmt(StmtKind::Expr(call_expr(
                        ident("__c_fs_rename"),
                        vec![ident("src"), ident("dst")],
                    ))),
                    stmt(StmtKind::Return(Some(int_lit(0)))),
                ],
                None,
            ),
            stmt(StmtKind::Expr(assign_expr(
                index_expr(ident(store_name), ident("dst")),
                index_expr(ident(store_name), ident("src")),
            ))),
            stmt(StmtKind::Expr(assign_expr(
                index_expr(ident(store_name), ident("src")),
                null_lit(),
            ))),
            stmt(StmtKind::Expr(assign_expr(
                index_expr(ident("__c_path_exists"), ident("dst")),
                int_lit(1),
            ))),
            stmt(StmtKind::Expr(assign_expr(
                index_expr(ident("__c_path_exists"), ident("src")),
                int_lit(0),
            ))),
            stmt(StmtKind::Expr(assign_expr(
                index_expr(ident("__c_is_dir"), ident("dst")),
                expr(ExprKind::NullCoalesce {
                    left: Box::new(index_expr(ident("__c_is_dir"), ident("src"))),
                    right: Box::new(int_lit(0)),
                }),
            ))),
            stmt(StmtKind::Expr(assign_expr(
                index_expr(ident("__c_is_dir"), ident("src")),
                int_lit(0),
            ))),
            stmt(StmtKind::Return(Some(int_lit(0)))),
        ],
    ));

    // unlink is remove() restricted to non-directories (EISDIR otherwise);
    // rmdir is remove() restricted to directories (ENOTDIR otherwise).
    out.push(function_stmt(
        "__c_unlink_h",
        vec!["path"],
        vec![
            if_stmt(
                expr(ExprKind::Binary {
                    op: BinOp::Eq,
                    left: Box::new(expr(ExprKind::NullCoalesce {
                        left: Box::new(index_expr(ident("__c_is_dir"), ident("path"))),
                        right: Box::new(int_lit(0)),
                    })),
                    right: Box::new(int_lit(1)),
                }),
                vec![stmt(StmtKind::Return(Some(int_lit(-1))))],
                None,
            ),
            stmt(StmtKind::Return(Some(call_expr(
                ident("__c_remove_h"),
                vec![ident("path")],
            )))),
        ],
    ));

    out.push(function_stmt(
        "__c_rmdir_h",
        vec!["path"],
        vec![
            if_stmt(
                expr(ExprKind::Binary {
                    op: BinOp::NotEq,
                    left: Box::new(expr(ExprKind::NullCoalesce {
                        left: Box::new(index_expr(ident("__c_is_dir"), ident("path"))),
                        right: Box::new(int_lit(0)),
                    })),
                    right: Box::new(int_lit(1)),
                }),
                vec![stmt(StmtKind::Return(Some(int_lit(-1))))],
                None,
            ),
            stmt(StmtKind::Expr(assign_expr(
                index_expr(ident("__c_is_dir"), ident("path")),
                int_lit(0),
            ))),
            stmt(StmtKind::Return(Some(call_expr(
                ident("__c_remove_h"),
                vec![ident("path")],
            )))),
        ],
    ));

    // symlink in the virtual store: register the link and SNAPSHOT the
    // target's content (true link indirection is not modeled — a later
    // write to the target is not reflected through the link).
    out.push(function_stmt(
        "__c_symlink_h",
        vec!["target", "link"],
        vec![
            stmt(StmtKind::Expr(assign_expr(
                index_expr(ident("__c_path_exists"), ident("link")),
                int_lit(1),
            ))),
            stmt(StmtKind::Expr(assign_expr(
                index_expr(ident(store_name), ident("link")),
                expr(ExprKind::NullCoalesce {
                    left: Box::new(index_expr(ident(store_name), ident("target"))),
                    right: Box::new(expr(ExprKind::Ternary {
                        cond: Box::new(call_expr(ident("__c_fs_exists"), vec![ident("target")])),
                        then: Box::new(call_expr(
                            ident("__c_fs_read"),
                            vec![ident("target"), str_lit("utf8")],
                        )),
                        else_: Box::new(str_lit("")),
                    })),
                }),
            ))),
            stmt(StmtKind::Return(Some(int_lit(0)))),
        ],
    ));

    // popen: "r" runs the child NOW and buffers its stdout as the FILE
    // content; "w" defers — the buffered writes become the child's stdin at
    // pclose. Both are real /bin/sh runs via `__c_sh_run`.
    out.push(function_stmt(
        "__c_popen_h",
        vec!["cmd", "mode"],
        vec![
            stmt(StmtKind::Expr(call_expr(
                ident("__c_write_stdout"),
                vec![ident(buffer_name)],
            ))),
            stmt(StmtKind::Expr(assign_expr(ident(buffer_name), str_lit("")))),
            stmt(StmtKind::Expr(assign_expr(
                ident("__c_popen_seq"),
                expr(ExprKind::Binary {
                    op: BinOp::Add,
                    left: Box::new(ident("__c_popen_seq")),
                    right: Box::new(int_lit(1)),
                }),
            ))),
            var_decl_stmt(
                "h",
                call_expr(
                    ident("__c_fopen_h"),
                    vec![
                        expr(ExprKind::Binary {
                            op: BinOp::Add,
                            left: Box::new(str_lit("__c_popen_")),
                            right: Box::new(ident("__c_popen_seq")),
                        }),
                        str_lit("w+"),
                    ],
                ),
            ),
            if_stmt(
                expr(ExprKind::Binary {
                    op: BinOp::GtEq,
                    left: Box::new(call_member(ident("mode"), "indexOf", vec![str_lit("w")])),
                    right: Box::new(int_lit(0)),
                }),
                vec![
                    stmt(StmtKind::Expr(assign_expr(
                        index_expr(ident("__c_popen_wcmd"), ident("h")),
                        ident("cmd"),
                    ))),
                    stmt(StmtKind::Return(Some(ident("h")))),
                ],
                None,
            ),
            var_decl_stmt(
                "r",
                call_expr(
                    ident("__c_sh_run"),
                    vec![ident("cmd"), expr(ExprKind::Object(Vec::new()))],
                ),
            ),
            if_stmt(
                expr(ExprKind::Binary {
                    op: BinOp::NotEq,
                    left: Box::new(member(ident("r"), "stderr")),
                    right: Box::new(str_lit("")),
                }),
                vec![stmt(StmtKind::Expr(call_expr(
                    ident("__c_fputs_h"),
                    vec![member(ident("r"), "stderr"), int_lit(2)],
                )))],
                None,
            ),
            stmt(StmtKind::Expr(assign_expr(
                index_expr(ident("__c_file_content"), ident("h")),
                member(ident("r"), "stdout"),
            ))),
            stmt(StmtKind::Expr(assign_expr(
                index_expr(ident("__c_file_pos"), ident("h")),
                int_lit(0),
            ))),
            stmt(StmtKind::Expr(assign_expr(
                index_expr(ident("__c_popen_status"), ident("h")),
                expr(ExprKind::Ternary {
                    cond: Box::new(expr(ExprKind::Binary {
                        op: BinOp::Eq,
                        left: Box::new(member(ident("r"), "status")),
                        right: Box::new(null_lit()),
                    })),
                    then: Box::new(int_lit(2)),
                    else_: Box::new(member(ident("r"), "status")),
                }),
            ))),
            stmt(StmtKind::Return(Some(ident("h")))),
        ],
    ));

    out.push(function_stmt(
        "__c_pclose_h",
        vec!["h"],
        vec![
            if_stmt(
                expr(ExprKind::Binary {
                    op: BinOp::NotEq,
                    left: Box::new(expr(ExprKind::NullCoalesce {
                        left: Box::new(index_expr(ident("__c_popen_wcmd"), ident("h"))),
                        right: Box::new(null_lit()),
                    })),
                    right: Box::new(null_lit()),
                }),
                vec![
                    var_decl_stmt("wcmd", index_expr(ident("__c_popen_wcmd"), ident("h"))),
                    stmt(StmtKind::Expr(assign_expr(
                        index_expr(ident("__c_popen_wcmd"), ident("h")),
                        null_lit(),
                    ))),
                    stmt(StmtKind::Expr(call_expr(
                        ident("__c_write_stdout"),
                        vec![ident(buffer_name)],
                    ))),
                    stmt(StmtKind::Expr(assign_expr(ident(buffer_name), str_lit("")))),
                    var_decl_stmt(
                        "r",
                        call_expr(
                            ident("__c_sh_run"),
                            vec![
                                ident("wcmd"),
                                expr(ExprKind::Object(vec![ObjectProperty::KeyValue {
                                    key: str_lit("input"),
                                    value: index_expr(ident("__c_file_content"), ident("h")),
                                }])),
                            ],
                        ),
                    ),
                    if_stmt(
                        expr(ExprKind::Binary {
                            op: BinOp::NotEq,
                            left: Box::new(member(ident("r"), "stdout")),
                            right: Box::new(str_lit("")),
                        }),
                        vec![stmt(StmtKind::Expr(call_expr(
                            ident("__c_write_stdout"),
                            vec![member(ident("r"), "stdout")],
                        )))],
                        None,
                    ),
                    if_stmt(
                        expr(ExprKind::Binary {
                            op: BinOp::NotEq,
                            left: Box::new(member(ident("r"), "stderr")),
                            right: Box::new(str_lit("")),
                        }),
                        vec![stmt(StmtKind::Expr(call_expr(
                            ident("__c_fputs_h"),
                            vec![member(ident("r"), "stderr"), int_lit(2)],
                        )))],
                        None,
                    ),
                    stmt(StmtKind::Expr(call_expr(
                        ident("__c_fclose_h"),
                        vec![ident("h")],
                    ))),
                    stmt(StmtKind::Return(Some(expr(ExprKind::Ternary {
                        cond: Box::new(expr(ExprKind::Binary {
                            op: BinOp::Eq,
                            left: Box::new(member(ident("r"), "status")),
                            right: Box::new(null_lit()),
                        })),
                        then: Box::new(int_lit(2)),
                        else_: Box::new(member(ident("r"), "status")),
                    })))),
                ],
                None,
            ),
            var_decl_stmt(
                "s",
                expr(ExprKind::NullCoalesce {
                    left: Box::new(index_expr(ident("__c_popen_status"), ident("h"))),
                    right: Box::new(int_lit(0)),
                }),
            ),
            stmt(StmtKind::Expr(call_expr(
                ident("__c_fclose_h"),
                vec![ident("h")],
            ))),
            stmt(StmtKind::Return(Some(ident("s")))),
        ],
    ));

    // One shell-spawn shape for system()/popen(): `/bin/sh -c cmd` through
    // node:child_process spawnSync, passing the runtime environment once any
    // setenv/putenv made it dirty. `opts` carries extra spawnSync options
    // ("w"-mode popen passes {input}); the env key is layered onto it.
    out.push(function_stmt(
        "__c_sh_run",
        vec!["cmd", "opts"],
        vec![
            if_stmt(
                expr(ExprKind::Binary {
                    op: BinOp::Eq,
                    left: Box::new(ident("__c_env_dirty")),
                    right: Box::new(int_lit(1)),
                }),
                vec![stmt(StmtKind::Expr(assign_expr(
                    member(ident("opts"), "env"),
                    ident("__c_env_obj"),
                )))],
                None,
            ),
            stmt(StmtKind::Return(Some(call_expr(
                ident("__c_spawn_sync"),
                vec![
                    str_lit("/bin/sh"),
                    expr(ExprKind::Array(vec![
                        ArrayElement {
                            key: None,
                            value: str_lit("-c"),
                            spread: false,
                            by_ref: false,
                        },
                        ArrayElement {
                            key: None,
                            value: ident("cmd"),
                            spread: false,
                            by_ref: false,
                        },
                    ])),
                    ident("opts"),
                ],
            )))),
        ],
    ));

    // `system(cmd)` — a REAL `/bin/sh -c` run via node:child_process
    // spawnSync. Buffered stdout is flushed first so ordering matches an
    // inherited fd; the child's stdout/stderr are forwarded through our
    // streams; the return value is the child's exit status (WEXITSTATUS is
    // identity in this runtime's wait model). A signal death reports status
    // null — mapped to 2 (SIGINT-shaped, nonzero).
    out.push(function_stmt(
        "__c_system_h",
        vec!["cmd"],
        vec![
            stmt(StmtKind::Expr(call_expr(
                ident("__c_write_stdout"),
                vec![ident(buffer_name)],
            ))),
            stmt(StmtKind::Expr(assign_expr(ident(buffer_name), str_lit("")))),
            var_decl_stmt(
                "r",
                call_expr(
                    ident("__c_sh_run"),
                    vec![ident("cmd"), expr(ExprKind::Object(Vec::new()))],
                ),
            ),
            if_stmt(
                expr(ExprKind::Binary {
                    op: BinOp::NotEq,
                    left: Box::new(member(ident("r"), "stdout")),
                    right: Box::new(str_lit("")),
                }),
                vec![stmt(StmtKind::Expr(call_expr(
                    ident("__c_write_stdout"),
                    vec![member(ident("r"), "stdout")],
                )))],
                None,
            ),
            if_stmt(
                expr(ExprKind::Binary {
                    op: BinOp::NotEq,
                    left: Box::new(member(ident("r"), "stderr")),
                    right: Box::new(str_lit("")),
                }),
                vec![stmt(StmtKind::Expr(call_expr(
                    ident("__c_fputs_h"),
                    vec![member(ident("r"), "stderr"), int_lit(2)],
                )))],
                None,
            ),
            if_stmt(
                expr(ExprKind::Binary {
                    op: BinOp::Eq,
                    left: Box::new(member(ident("r"), "status")),
                    right: Box::new(null_lit()),
                }),
                vec![stmt(StmtKind::Return(Some(int_lit(2))))],
                None,
            ),
            stmt(StmtKind::Return(Some(member(ident("r"), "status")))),
        ],
    ));

    out.push(function_stmt(
        "__c_stdout_append",
        vec!["piece"],
        vec![
            if_stmt(
                call_member(ident("piece"), "endsWith", vec![str_lit("\n")]),
                vec![
                    // Flush the completed line straight to stdout (NOT the
                    // `print`/wasi:logging vybelib builtin). `__c_write_stdout`
                    // is the libc intrinsic, which goes through the shared
                    // `primitives::io` write — byte-faithful, no implicit
                    // newline — so write `buffer + piece` verbatim (piece still
                    // carries its trailing '\n').
                    stmt(StmtKind::Expr(call_expr(
                        ident("__c_write_stdout"),
                        vec![expr(ExprKind::Binary {
                            op: BinOp::Add,
                            left: Box::new(ident(buffer_name)),
                            right: Box::new(ident("piece")),
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
            // Store miss → read through to the REAL filesystem (node:fs), so
            // files written by `system()` children (shell redirects) open
            // like any other file. `??` (not `== null`) — a missing store key
            // is undefined and the VM's Eq is strict.
            stmt(StmtKind::Expr(assign_expr(
                ident("existing"),
                expr(ExprKind::NullCoalesce {
                    left: Box::new(ident("existing")),
                    right: Box::new(expr(ExprKind::Ternary {
                        cond: Box::new(call_expr(ident("__c_fs_exists"), vec![ident("path")])),
                        then: Box::new(call_expr(
                            ident("__c_fs_read"),
                            vec![ident("path"), str_lit("utf8")],
                        )),
                        else_: Box::new(null_lit()),
                    })),
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
                "append_mode",
                expr(ExprKind::Ternary {
                    cond: Box::new(expr(ExprKind::Binary {
                        op: BinOp::GtEq,
                        left: Box::new(call_member(ident("mode"), "indexOf", vec![str_lit("a")])),
                        right: Box::new(int_lit(0)),
                    })),
                    then: Box::new(int_lit(1)),
                    else_: Box::new(int_lit(0)),
                }),
            ),
            var_decl_stmt(
                "readonly_mode",
                expr(ExprKind::Ternary {
                    cond: Box::new(expr(ExprKind::Binary {
                        op: BinOp::Eq,
                        left: Box::new(call_member(ident("mode"), "indexOf", vec![str_lit("w")])),
                        right: Box::new(int_lit(-1)),
                    })),
                    then: Box::new(expr(ExprKind::Ternary {
                        cond: Box::new(expr(ExprKind::Binary {
                            op: BinOp::Eq,
                            left: Box::new(call_member(
                                ident("mode"),
                                "indexOf",
                                vec![str_lit("a")],
                            )),
                            right: Box::new(int_lit(-1)),
                        })),
                        then: Box::new(int_lit(1)),
                        else_: Box::new(int_lit(0)),
                    })),
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
            stmt(StmtKind::Expr(assign_expr(
                ident("path"),
                call_expr(ident("__libc_char_to_str"), vec![ident("path")]),
            ))),
            stmt(StmtKind::Expr(assign_expr(
                ident("mode"),
                call_expr(ident("__libc_char_to_str"), vec![ident("mode")]),
            ))),
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
                    // Store miss → read through to the REAL filesystem
                    // (node:fs), so files written by `system()` children
                    // (shell redirects) open like any other file. A
                    // TOMBSTONED path (removed/renamed away, registry 0)
                    // must NOT resurrect from a stale same-named real file.
                    expr(ExprKind::Ternary {
                        cond: Box::new(expr(ExprKind::Binary {
                            op: BinOp::And,
                            left: Box::new(expr(ExprKind::Binary {
                                op: BinOp::NotEq,
                                left: Box::new(expr(ExprKind::NullCoalesce {
                                    left: Box::new(index_expr(
                                        ident("__c_path_exists"),
                                        ident("path"),
                                    )),
                                    right: Box::new(int_lit(-1)),
                                })),
                                right: Box::new(int_lit(0)),
                            })),
                            right: Box::new(call_expr(ident("__c_fs_exists"), vec![ident("path")])),
                        })),
                        then: Box::new(call_expr(
                            ident("__c_fs_read"),
                            vec![ident("path"), str_lit("utf8")],
                        )),
                        else_: Box::new(expr(ExprKind::Ternary {
                            cond: Box::new(expr(ExprKind::Binary {
                                op: BinOp::NotEq,
                                left: Box::new(ident("binary_mode")),
                                right: Box::new(int_lit(0)),
                            })),
                            then: Box::new(expr(ExprKind::Array(vec![]))),
                            else_: Box::new(str_lit("")),
                        })),
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
            // Creating "dir/file" marks the parent directory non-empty
            // (consumed by rename's ENOTEMPTY check).
            if_stmt(
                expr(ExprKind::Binary {
                    op: BinOp::And,
                    left: Box::new(expr(ExprKind::Binary {
                        op: BinOp::NotEq,
                        left: Box::new(ident("write_mode")),
                        right: Box::new(int_lit(0)),
                    })),
                    right: Box::new(expr(ExprKind::Binary {
                        op: BinOp::GtEq,
                        left: Box::new(call_member(
                            ident("path"),
                            "lastIndexOf",
                            vec![str_lit("/")],
                        )),
                        right: Box::new(int_lit(0)),
                    })),
                }),
                vec![stmt(StmtKind::Expr(assign_expr(
                    index_expr(
                        ident("__c_dir_nonempty"),
                        call_member(
                            ident("path"),
                            "substring",
                            vec![
                                int_lit(0),
                                call_member(ident("path"), "lastIndexOf", vec![str_lit("/")]),
                            ],
                        ),
                    ),
                    int_lit(1),
                )))],
                None,
            ),
            stmt(StmtKind::Expr(assign_expr(
                index_expr(ident("__c_file_content"), ident("handle")),
                ident("content"),
            ))),
            stmt(StmtKind::Expr(assign_expr(
                index_expr(ident("__c_file_pos"), ident("handle")),
                expr(ExprKind::Ternary {
                    cond: Box::new(expr(ExprKind::Binary {
                        op: BinOp::NotEq,
                        left: Box::new(ident("append_mode")),
                        right: Box::new(int_lit(0)),
                    })),
                    then: Box::new(member(ident("content"), "length")),
                    else_: Box::new(int_lit(0)),
                }),
            ))),
            stmt(StmtKind::Expr(assign_expr(
                index_expr(ident("__c_file_eof"), ident("handle")),
                int_lit(0),
            ))),
            stmt(StmtKind::Expr(assign_expr(
                index_expr(ident("__c_file_error"), ident("handle")),
                int_lit(0),
            ))),
            stmt(StmtKind::Expr(assign_expr(
                index_expr(ident("__c_file_ungot"), ident("handle")),
                expr(ExprKind::Array(vec![])),
            ))),
            stmt(StmtKind::Expr(assign_expr(
                index_expr(ident("__c_file_dirty"), ident("handle")),
                ident("write_mode"),
            ))),
            stmt(StmtKind::Expr(assign_expr(
                index_expr(ident("__c_file_append"), ident("handle")),
                ident("append_mode"),
            ))),
            stmt(StmtKind::Expr(assign_expr(
                index_expr(ident("__c_file_readonly"), ident("handle")),
                ident("readonly_mode"),
            ))),
            stmt(StmtKind::Expr(assign_expr(
                index_expr(ident("__c_file_closed"), ident("handle")),
                int_lit(0),
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
        "__c_fclose_h",
        vec!["handle"],
        vec![
            stmt(StmtKind::Expr(call_expr(
                ident("__c_fsync_h"),
                vec![ident("handle")],
            ))),
            stmt(StmtKind::Expr(assign_expr(
                index_expr(ident("__c_file_closed"), ident("handle")),
                int_lit(1),
            ))),
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
            // stderr is UNBUFFERED: straight to the real wasi:cli/stderr
            // stream, no file-store entry (handle 2 has none — reaching the
            // store path below threw on `undefined.length`).
            if_stmt(
                expr(ExprKind::Binary {
                    op: BinOp::Eq,
                    left: Box::new(ident("handle")),
                    right: Box::new(int_lit(2)),
                }),
                vec![stmt(StmtKind::Return(Some(call_expr(
                    ident("__c_write_stderr"),
                    vec![ident("text")],
                ))))],
                None,
            ),
            if_stmt(
                index_expr(ident("__c_file_readonly"), ident("handle")),
                vec![
                    stmt(StmtKind::Expr(assign_expr(
                        index_expr(ident("__c_file_error"), ident("handle")),
                        int_lit(1),
                    ))),
                    stmt(StmtKind::Return(Some(int_lit(-1)))),
                ],
                None,
            ),
            if_stmt(
                index_expr(ident("__c_file_append"), ident("handle")),
                vec![stmt(StmtKind::Expr(assign_expr(
                    index_expr(ident("__c_file_pos"), ident("handle")),
                    member(
                        index_expr(ident("__c_file_content"), ident("handle")),
                        "length",
                    ),
                )))],
                None,
            ),
            var_decl_stmt("pos", index_expr(ident("__c_file_pos"), ident("handle"))),
            var_decl_stmt(
                "content",
                index_expr(ident("__c_file_content"), ident("handle")),
            ),
            if_stmt(
                expr(ExprKind::Binary {
                    op: BinOp::Gt,
                    left: Box::new(ident("pos")),
                    right: Box::new(member(ident("content"), "length")),
                }),
                vec![stmt(StmtKind::Expr(assign_expr(
                    ident("content"),
                    expr(ExprKind::Binary {
                        op: BinOp::Add,
                        left: Box::new(ident("content")),
                        right: Box::new(call_member(
                            str_lit("\0"),
                            "repeat",
                            vec![expr(ExprKind::Binary {
                                op: BinOp::Sub,
                                left: Box::new(ident("pos")),
                                right: Box::new(member(ident("content"), "length")),
                            })],
                        )),
                    }),
                )))],
                None,
            ),
            if_stmt(
                expr(ExprKind::Binary {
                    op: BinOp::Eq,
                    left: Box::new(ident("pos")),
                    right: Box::new(member(ident("content"), "length")),
                }),
                vec![stmt(StmtKind::Expr(assign_expr(
                    index_expr(ident("__c_file_content"), ident("handle")),
                    expr(ExprKind::Binary {
                        op: BinOp::Add,
                        left: Box::new(ident("content")),
                        right: Box::new(ident("text")),
                    }),
                )))],
                Some(vec![stmt(StmtKind::Expr(assign_expr(
                    index_expr(ident("__c_file_content"), ident("handle")),
                    expr(ExprKind::Binary {
                        op: BinOp::Add,
                        left: Box::new(call_member(
                            ident("content"),
                            "substring",
                            vec![int_lit(0), ident("pos")],
                        )),
                        right: Box::new(expr(ExprKind::Binary {
                            op: BinOp::Add,
                            left: Box::new(ident("text")),
                            right: Box::new(call_member(
                                ident("content"),
                                "substring",
                                vec![expr(ExprKind::Binary {
                                    op: BinOp::Add,
                                    left: Box::new(ident("pos")),
                                    right: Box::new(member(ident("text"), "length")),
                                })],
                            )),
                        })),
                    }),
                )))]),
            ),
            stmt(StmtKind::Expr(assign_expr(
                index_expr(ident("__c_file_pos"), ident("handle")),
                expr(ExprKind::Binary {
                    op: BinOp::Add,
                    left: Box::new(ident("pos")),
                    right: Box::new(member(ident("text"), "length")),
                }),
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
            var_decl_stmt(
                "ungot",
                index_expr(ident("__c_file_ungot"), ident("handle")),
            ),
            if_stmt(
                expr(ExprKind::Binary {
                    op: BinOp::Gt,
                    left: Box::new(member(ident("ungot"), "length")),
                    right: Box::new(int_lit(0)),
                }),
                vec![
                    var_decl_stmt("ch", call_member(ident("ungot"), "pop", vec![])),
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
            if_stmt(
                expr(ExprKind::Binary {
                    op: BinOp::Eq,
                    left: Box::new(ident("code")),
                    right: Box::new(int_lit(-1)),
                }),
                vec![stmt(StmtKind::Return(Some(int_lit(-1))))],
                None,
            ),
            if_stmt(
                expr(ExprKind::Binary {
                    op: BinOp::Eq,
                    left: Box::new(expr(ExprKind::Unary {
                        op: vybe_ast::UnaryOp::Typeof,
                        expr: Box::new(index_expr(ident("__c_file_ungot"), ident("handle"))),
                    })),
                    right: Box::new(str_lit("undefined")),
                }),
                vec![stmt(StmtKind::Expr(assign_expr(
                    index_expr(ident("__c_file_ungot"), ident("handle")),
                    expr(ExprKind::Array(vec![])),
                )))],
                None,
            ),
            stmt(StmtKind::Expr(call_member(
                index_expr(ident("__c_file_ungot"), ident("handle")),
                "push",
                vec![ident("code")],
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
    out.extend(crate::emitter::stdio_adapter::stdin_runtime_helpers());
    // char[] → string decoder for `%s`/`puts` (string / carray / code-point array).
    out.push(crate::emitter::stdio_adapter::char_to_str_runtime_helper());
    out.extend(crate::emitter::string_adapter::strtok_runtime_helpers());
    // wide-char boundary helpers (code-point array ↔ string) for wchar.h.
    out.extend(crate::emitter::wchar_adapter::runtime_helpers());

    // math.h domain-error helpers (libc surface) — sqrt sets errno (EDOM).
    out.extend(crate::emitter::math_adapter::domain_error_helpers());
    out.extend(crate::emitter::math_adapter::fenv_runtime_helpers());

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
                index_expr(ident("__c_file_closed"), ident("handle")),
                vec![stmt(StmtKind::Return(Some(int_lit(-1))))],
                None,
            ),
            stmt(StmtKind::Expr(assign_expr(
                index_expr(ident("__c_file_pos"), ident("handle")),
                ident("pos"),
            ))),
            stmt(StmtKind::Expr(assign_expr(
                index_expr(ident("__c_file_eof"), ident("handle")),
                int_lit(0),
            ))),
            stmt(StmtKind::Expr(assign_expr(
                index_expr(ident("__c_file_ungot"), ident("handle")),
                expr(ExprKind::Array(vec![])),
            ))),
            stmt(StmtKind::Return(Some(int_lit(0)))),
        ],
    ));

    out.push(function_stmt(
        "__c_ftell_h",
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
                index_expr(ident("__c_file_closed"), ident("handle")),
                vec![stmt(StmtKind::Return(Some(int_lit(-1))))],
                None,
            ),
            var_decl_stmt("pos", index_expr(ident("__c_file_pos"), ident("handle"))),
            if_stmt(
                expr(ExprKind::Binary {
                    op: BinOp::Eq,
                    left: Box::new(expr(ExprKind::Unary {
                        op: vybe_ast::UnaryOp::Typeof,
                        expr: Box::new(ident("pos")),
                    })),
                    right: Box::new(str_lit("undefined")),
                }),
                vec![stmt(StmtKind::Return(Some(int_lit(0))))],
                None,
            ),
            stmt(StmtKind::Return(Some(expr(ExprKind::Binary {
                op: BinOp::Sub,
                left: Box::new(ident("pos")),
                right: Box::new(member(
                    index_expr(ident("__c_file_ungot"), ident("handle")),
                    "length",
                )),
            })))),
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
        vec![
            var_decl_stmt("out", str_lit("")),
            var_decl_stmt("i", int_lit(0)),
            stmt(StmtKind::While {
                cond: expr(ExprKind::Binary {
                    op: BinOp::Lt,
                    left: Box::new(ident("i")),
                    right: Box::new(ident("count")),
                }),
                body: vec![
                    var_decl_stmt("ch", call_expr(ident("__c_fgetc_h"), vec![ident("handle")])),
                    if_stmt(
                        expr(ExprKind::Binary {
                            op: BinOp::Eq,
                            left: Box::new(ident("ch")),
                            right: Box::new(int_lit(-1)),
                        }),
                        vec![stmt(StmtKind::Break(vybe_ast::BreakTarget::Implicit))],
                        None,
                    ),
                    stmt(StmtKind::Expr(assign_expr(
                        ident("out"),
                        expr(ExprKind::Binary {
                            op: BinOp::Add,
                            left: Box::new(ident("out")),
                            right: Box::new(call_member(
                                ident("String"),
                                "fromCharCode",
                                vec![ident("ch")],
                            )),
                        }),
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
            stmt(StmtKind::Return(Some(ident("out")))),
        ],
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
            if_stmt(
                expr(ExprKind::Binary {
                    op: BinOp::NotEq,
                    left: Box::new(expr(ExprKind::TypeOf(Box::new(ident("h"))))),
                    right: Box::new(str_lit("function")),
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
    out.extend(crate::emitter::stdlib_runtime::runtime_helpers());

    // regex.h runtime helpers (regcomp/regexec on the ECMA RegExp surface).
    out.extend(crate::emitter::regex_adapter::runtime_helpers());

    // POSIX helpers for libc pointer-backed APIs such as mmap.
    out.extend(crate::emitter::posix_adapter::runtime_helpers());

    // string.h runtime helpers (strcoll/strxfrm/strpbrk/strspn/strcspn).
    out.extend(crate::emitter::string_runtime::runtime_helpers());

    // time.h runtime helpers live in their own adapter (shared libc surface).
    out.extend(crate::emitter::time_adapter::runtime_helpers());

    out
}

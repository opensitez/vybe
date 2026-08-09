//! pthreads, POSIX semaphores, and C11 threads adapters for libc.

use vybe_ast::{BinOp, ExprKind, Expression, Literal, PlaceExpr, Statement, StmtKind, UnaryOp};

use super::build::{
    assign_expr, expr, function_stmt, ident, if_stmt, index_expr, int_lit, member, null_lit, stmt,
    var_decl_stmt,
};

pub fn header_constants(header: &str) -> Option<&'static [(&'static str, i64)]> {
    match header {
        "pthread.h" => Some(&[
            ("PTHREAD_MUTEX_INITIALIZER", 0),
            ("PTHREAD_RECURSIVE_MUTEX_INITIALIZER_NP", 20),
            ("PTHREAD_ERRORCHECK_MUTEX_INITIALIZER_NP", 10),
            ("PTHREAD_MUTEX_NORMAL", 0),
            ("PTHREAD_MUTEX_ERRORCHECK", 1),
            ("PTHREAD_MUTEX_RECURSIVE", 2),
            ("PTHREAD_COND_INITIALIZER", 0),
            ("PTHREAD_RWLOCK_INITIALIZER", 0),
            ("CLOCK_REALTIME", 0),
            ("CLOCK_MONOTONIC", 1),
            ("PTHREAD_PROCESS_PRIVATE", 0),
            ("PTHREAD_PROCESS_SHARED", 1),
            ("PTHREAD_PRIO_NONE", 0),
            ("PTHREAD_PRIO_INHERIT", 1),
            ("PTHREAD_PRIO_PROTECT", 2),
            ("PTHREAD_MUTEX_STALLED", 0),
            ("PTHREAD_MUTEX_ROBUST", 1),
            ("PTHREAD_ONCE_INIT", 0),
            ("PTHREAD_CREATE_JOINABLE", 0),
            ("PTHREAD_CREATE_DETACHED", 1),
            ("PTHREAD_CANCEL_ENABLE", 0),
            ("PTHREAD_CANCEL_DISABLE", 1),
            ("PTHREAD_CANCEL_DEFERRED", 0),
            ("PTHREAD_CANCEL_ASYNCHRONOUS", 1),
            ("PTHREAD_CANCELED", -1),
            ("PTHREAD_BARRIER_SERIAL_THREAD", 1),
        ]),
        "semaphore.h" => Some(&[("SEM_FAILED", -1), ("SEM_VALUE_MAX", 32767)]),
        "threads.h" => Some(&[
            ("thrd_success", 0),
            ("thrd_error", 1),
            ("thrd_busy", 2),
            ("thrd_nomem", 3),
            ("thrd_timedout", 4),
            ("mtx_plain", 0),
            ("mtx_recursive", 1),
            ("mtx_timed", 2),
            ("ONCE_FLAG_INIT", 0),
        ]),
        _ => None,
    }
}

pub fn ok() -> Expression {
    int_lit(0)
}

pub fn fail() -> Expression {
    int_lit(-1)
}

pub fn init_target(target: Expression, value: Expression) -> Expression {
    expr(ExprKind::Sequence(vec![
        assign_expr(arg_target(target), value),
        int_lit(0),
    ]))
}

/// `pthread_create` / `thrd_create` — a REAL thread.
///
/// The spawn itself is the shared lowering (`common:threading.thread_spawn`
/// → funcref `table.grow`, a 16-byte `{fn_idx, status, user_arg}` record in
/// the shared futex page, then the `wasi:threads.thread-spawn` import). Only
/// the C-side bookkeeping lives here: a `pthread_t` is an integer handle, so
/// the handle — not the `void*` — is what rides the record's i32 `user_arg`
/// word. The start function and its argument travel in `__c_thread_starts` /
/// `__c_thread_args`, plain objects, which the child sees through the same
/// `Arc` its globals clone carries. `__c_thread_entry` unpacks them.
pub fn pthread_create(thread: Expression, start: Expression, arg: Expression) -> Expression {
    let target = arg_target(thread);
    expr(ExprKind::Sequence(vec![
        assign_expr(target.clone(), ident("__c_next_thread_handle")),
        assign_expr(
            ident("__c_next_thread_handle"),
            add(ident("__c_next_thread_handle"), int_lit(1)),
        ),
        // Published BEFORE the spawn — the child reads both out of the
        // shared tables the moment it starts.
        assign_expr(
            index_expr(ident("__c_thread_starts"), target.clone()),
            start,
        ),
        assign_expr(index_expr(ident("__c_thread_args"), target.clone()), arg),
        assign_expr(
            index_expr(ident("__c_thread_results"), target.clone()),
            null_lit(),
        ),
        assign_expr(
            index_expr(ident("__c_thread_tasks"), target.clone()),
            expr(ExprKind::Call {
                callee: Box::new(ident("__c_thread_spawn")),
                args: vec![
                    vybe_ast::Argument::positional(target),
                    vybe_ast::Argument::positional(ident("__c_thread_entry")),
                ],
                optional: false,
            }),
        ),
        int_lit(0),
    ]))
}

pub fn thrd_create(thread: Expression, start: Expression, arg: Expression) -> Expression {
    pthread_create(thread, start, arg)
}

/// `pthread_join` / `thrd_join` — block on the spawned task, then report the
/// start function's return value. The wait is `common:threading.thread_join`
/// (a futex wait on the record's status word); the value comes back through
/// `__c_thread_results`, since the status word only carries done/faulted.
pub fn pthread_join(thread: Expression, retval: Expression) -> Expression {
    let joined = expr(ExprKind::Call {
        callee: Box::new(ident("__c_thread_join_h")),
        args: vec![vybe_ast::Argument::positional(thread)],
        optional: false,
    });
    if matches!(retval.kind, ExprKind::Lit(Literal::Null)) {
        return expr(ExprKind::Sequence(vec![joined, int_lit(0)]));
    }
    expr(ExprKind::Sequence(vec![
        assign_expr(arg_target(retval), joined),
        int_lit(0),
    ]))
}

pub fn pthread_cancel(thread: Expression) -> Expression {
    expr(ExprKind::Sequence(vec![
        assign_expr(index_expr(ident("__c_thread_results"), thread), int_lit(-1)),
        int_lit(0),
    ]))
}

pub fn equal(left: Expression, right: Expression) -> Expression {
    expr(ExprKind::Ternary {
        cond: Box::new(eq(left, right)),
        then: Box::new(int_lit(1)),
        else_: Box::new(int_lit(0)),
    })
}

pub fn call_once(flag: Expression, init: Expression) -> Expression {
    expr(ExprKind::Ternary {
        cond: Box::new(eq(arg_target(flag.clone()), int_lit(1))),
        then: Box::new(int_lit(0)),
        else_: Box::new(expr(ExprKind::Sequence(vec![
            assign_expr(arg_target(flag), int_lit(1)),
            expr(ExprKind::Call {
                callee: Box::new(init),
                args: vec![],
                optional: false,
            }),
            int_lit(0),
        ]))),
    })
}

pub fn cleanup_push(cleanup: Expression, arg: Expression) -> Expression {
    expr(ExprKind::Sequence(vec![
        assign_expr(ident("__c_cleanup_fn"), cleanup),
        assign_expr(ident("__c_cleanup_arg"), arg),
        int_lit(0),
    ]))
}

pub fn cleanup_pop(execute: Expression) -> Expression {
    expr(ExprKind::Ternary {
        cond: Box::new(execute),
        then: Box::new(expr(ExprKind::Sequence(vec![
            expr(ExprKind::Call {
                callee: Box::new(ident("__c_cleanup_fn")),
                args: vec![vybe_ast::Argument::positional(ident("__c_cleanup_arg"))],
                optional: false,
            }),
            int_lit(0),
        ]))),
        else_: Box::new(int_lit(0)),
    })
}

pub fn mutex_init(mutex: Expression, attr: Option<Expression>) -> Expression {
    let attr_value = attr.map_or_else(|| int_lit(0), attr_target);
    let value = expr(ExprKind::Ternary {
        cond: Box::new(eq(attr_value.clone(), int_lit(1))),
        then: Box::new(int_lit(10)),
        else_: Box::new(expr(ExprKind::Ternary {
            cond: Box::new(eq(attr_value, int_lit(2))),
            then: Box::new(int_lit(20)),
            else_: Box::new(int_lit(0)),
        })),
    });
    init_target(mutex, value)
}

pub fn mutex_lock(mutex: Expression) -> Expression {
    let target = arg_target(mutex);
    expr(ExprKind::Ternary {
        cond: Box::new(eq(target.clone(), int_lit(11))),
        then: Box::new(int_lit(16)),
        else_: Box::new(expr(ExprKind::Sequence(vec![
            assign_expr(
                target.clone(),
                expr(ExprKind::Ternary {
                    cond: Box::new(eq(target.clone(), int_lit(10))),
                    then: Box::new(int_lit(11)),
                    else_: Box::new(expr(ExprKind::Ternary {
                        cond: Box::new(expr(ExprKind::Binary {
                            op: BinOp::GtEq,
                            left: Box::new(target.clone()),
                            right: Box::new(int_lit(20)),
                        })),
                        then: Box::new(add(target, int_lit(1))),
                        else_: Box::new(int_lit(1)),
                    })),
                }),
            ),
            int_lit(0),
        ]))),
    })
}

pub fn mutex_trylock(mutex: Expression) -> Expression {
    let target = arg_target(mutex);
    expr(ExprKind::Ternary {
        cond: Box::new(or(
            eq(target.clone(), int_lit(1)),
            eq(target.clone(), int_lit(11)),
        )),
        then: Box::new(int_lit(16)),
        else_: Box::new(expr(ExprKind::Sequence(vec![
            assign_expr(
                target.clone(),
                expr(ExprKind::Ternary {
                    cond: Box::new(eq(target.clone(), int_lit(10))),
                    then: Box::new(int_lit(11)),
                    else_: Box::new(int_lit(1)),
                }),
            ),
            int_lit(0),
        ]))),
    })
}

pub fn mutex_timedlock(mutex: Expression) -> Expression {
    mutex_trylock(mutex)
}

pub fn mutex_unlock(mutex: Expression) -> Expression {
    let target = arg_target(mutex);
    expr(ExprKind::Ternary {
        cond: Box::new(eq(target.clone(), int_lit(10))),
        then: Box::new(int_lit(16)),
        else_: Box::new(expr(ExprKind::Sequence(vec![
            assign_expr(
                target.clone(),
                expr(ExprKind::Ternary {
                    cond: Box::new(eq(target.clone(), int_lit(11))),
                    then: Box::new(int_lit(10)),
                    else_: Box::new(expr(ExprKind::Ternary {
                        cond: Box::new(expr(ExprKind::Binary {
                            op: BinOp::Gt,
                            left: Box::new(target.clone()),
                            right: Box::new(int_lit(20)),
                        })),
                        then: Box::new(add(target, int_lit(-1))),
                        else_: Box::new(int_lit(0)),
                    })),
                }),
            ),
            int_lit(0),
        ]))),
    })
}

pub fn attr_set(attr: Expression, value: Expression) -> Expression {
    init_target(attr, value)
}

pub fn attr_get(attr: Expression, out: Expression, default_value: Expression) -> Expression {
    let value = expr(ExprKind::NullCoalesce {
        left: Box::new(arg_target(attr)),
        right: Box::new(default_value),
    });
    expr(ExprKind::Sequence(vec![
        assign_expr(arg_target(out), value),
        int_lit(0),
    ]))
}

pub fn rwlock_rdlock(lock: Expression) -> Expression {
    let target = arg_target(lock);
    expr(ExprKind::Sequence(vec![
        assign_expr(target.clone(), add(target, int_lit(1))),
        int_lit(0),
    ]))
}

pub fn rwlock_wrlock(lock: Expression) -> Expression {
    init_target(lock, int_lit(-1))
}

pub fn rwlock_tryrdlock(lock: Expression) -> Expression {
    let target = arg_target(lock);
    expr(ExprKind::Ternary {
        cond: Box::new(eq(target.clone(), int_lit(-1))),
        then: Box::new(int_lit(16)),
        else_: Box::new(expr(ExprKind::Sequence(vec![
            assign_expr(target.clone(), add(target, int_lit(1))),
            int_lit(0),
        ]))),
    })
}

pub fn rwlock_trywrlock(lock: Expression) -> Expression {
    let target = arg_target(lock);
    expr(ExprKind::Ternary {
        cond: Box::new(eq(target.clone(), int_lit(0))),
        then: Box::new(expr(ExprKind::Sequence(vec![
            assign_expr(target, int_lit(-1)),
            int_lit(0),
        ]))),
        else_: Box::new(int_lit(16)),
    })
}

pub fn rwlock_unlock(lock: Expression) -> Expression {
    init_target(lock, int_lit(0))
}

/// `pthread_barrier_init` — the barrier variable holds a HANDLE, not the
/// count. The count and the arrival tally live in shared objects instead,
/// because that is the only state a spawned thread can actually reach: a
/// child's globals are a clone (a scalar written there never comes back),
/// while an object is one `Arc` on both sides. The handle itself is written
/// before any spawn, so every child's clone names the same record.
pub fn barrier_init(barrier: Expression, count: Expression) -> Expression {
    let target = arg_target(barrier);
    expr(ExprKind::Ternary {
        cond: Box::new(lt(count.clone(), int_lit(1))),
        then: Box::new(int_lit(22)),
        else_: Box::new(expr(ExprKind::Sequence(vec![
            assign_expr(target.clone(), ident("__c_next_barrier_handle")),
            assign_expr(
                ident("__c_next_barrier_handle"),
                add(ident("__c_next_barrier_handle"), int_lit(1)),
            ),
            assign_expr(
                index_expr(ident("__c_barrier_limit"), target.clone()),
                count,
            ),
            assign_expr(
                index_expr(ident("__c_barrier_arrived"), target.clone()),
                int_lit(0),
            ),
            assign_expr(index_expr(ident("__c_barrier_gen"), target), int_lit(0)),
            int_lit(0),
        ]))),
    })
}

/// `pthread_barrier_wait` — arrive, then block until the generation turns
/// over. Real waiting on real threads; the previous version ran a pending
/// thread's BODY inline here, which with a real spawn would run it twice.
pub fn barrier_wait(barrier: Expression) -> Expression {
    expr(ExprKind::Call {
        callee: Box::new(ident("__c_barrier_wait_h")),
        args: vec![vybe_ast::Argument::positional(arg_target(barrier))],
        optional: false,
    })
}

pub fn timespec_get(ts: Expression, base: Expression) -> Expression {
    let target = arg_target(ts);
    expr(ExprKind::Sequence(vec![
        assign_expr(member(target.clone(), "tv_sec"), int_lit(0)),
        assign_expr(member(target, "tv_nsec"), int_lit(0)),
        base,
    ]))
}

pub fn gettimeofday(tv: Expression) -> Expression {
    let target = arg_target(tv);
    expr(ExprKind::Sequence(vec![
        assign_expr(member(target.clone(), "tv_sec"), int_lit(0)),
        assign_expr(member(target, "tv_usec"), int_lit(0)),
        int_lit(0),
    ]))
}

pub fn sem_open(name: Expression, flags: Expression, value: Option<Expression>) -> Expression {
    let exists = index_expr(ident("__c_sem_exists"), name.clone());
    let has_creat = bit_set(flags.clone(), 64);
    let has_excl = bit_set(flags.clone(), 128);
    let missing_without_creat = and(not(exists.clone()), not(has_creat));
    let exclusive_existing = and(exists.clone(), has_excl);
    let name_too_long = gt(member(name.clone(), "length"), int_lit(255));
    let initial = value.unwrap_or_else(|| int_lit(1));
    let handle = expr(ExprKind::NullCoalesce {
        left: Box::new(index_expr(ident("__c_sem_handles"), name.clone())),
        right: Box::new(ident("__c_next_sem_handle")),
    });
    let ok = expr(ExprKind::Sequence(vec![
        assign_expr(
            index_expr(ident("__c_sem_exists"), name.clone()),
            int_lit(1),
        ),
        assign_expr(
            index_expr(ident("__c_sem_handles"), name.clone()),
            handle.clone(),
        ),
        assign_expr(
            index_expr(ident("__c_sem_values"), handle.clone()),
            expr(ExprKind::NullCoalesce {
                left: Box::new(index_expr(ident("__c_sem_values"), handle.clone())),
                right: Box::new(initial),
            }),
        ),
        assign_expr(
            ident("__c_next_sem_handle"),
            add(ident("__c_next_sem_handle"), int_lit(1)),
        ),
        handle,
    ]));
    expr(ExprKind::Ternary {
        cond: Box::new(or(
            name_too_long,
            or(missing_without_creat, exclusive_existing),
        )),
        then: Box::new(int_lit(-1)),
        else_: Box::new(ok),
    })
}

pub fn sem_unlink(name: Expression) -> Expression {
    let exists = index_expr(ident("__c_sem_exists"), name.clone());
    expr(ExprKind::Ternary {
        cond: Box::new(exists),
        then: Box::new(expr(ExprKind::Sequence(vec![
            assign_expr(index_expr(ident("__c_sem_exists"), name), int_lit(0)),
            int_lit(0),
        ]))),
        else_: Box::new(int_lit(-1)),
    })
}

pub fn sem_init(sem: Expression, value: Expression) -> Expression {
    expr(ExprKind::Ternary {
        cond: Box::new(lt(value.clone(), int_lit(0))),
        then: Box::new(int_lit(-1)),
        else_: Box::new(init_target(sem, value)),
    })
}

pub fn sem_post(sem: Expression) -> Expression {
    let target = sem_target(sem.clone());
    expr(ExprKind::Sequence(vec![
        assign_expr(target, add(sem_value(sem), int_lit(1))),
        int_lit(0),
    ]))
}

pub fn sem_wait(sem: Expression, timed: bool, invalid_time: bool) -> Expression {
    let target = sem_target(sem.clone());
    let value = sem_value(sem);
    let fail_cond = if invalid_time {
        int_lit(1)
    } else {
        expr(ExprKind::Binary {
            op: BinOp::LtEq,
            left: Box::new(value.clone()),
            right: Box::new(int_lit(0)),
        })
    };
    expr(ExprKind::Ternary {
        cond: Box::new(fail_cond),
        then: Box::new(if timed { int_lit(-1) } else { int_lit(0) }),
        else_: Box::new(expr(ExprKind::Sequence(vec![
            assign_expr(target, add(value, int_lit(-1))),
            int_lit(0),
        ]))),
    })
}

pub fn sem_getvalue(sem: Expression, out: Expression) -> Expression {
    expr(ExprKind::Sequence(vec![
        assign_expr(arg_target(out), sem_value(sem)),
        int_lit(0),
    ]))
}

pub fn key_create(key: Expression) -> Expression {
    let target = arg_target(key);
    expr(ExprKind::Sequence(vec![
        assign_expr(target.clone(), ident("__c_next_tls_key")),
        assign_expr(ident("__c_last_tls_key"), target.clone()),
        assign_expr(
            ident("__c_next_tls_key"),
            add(ident("__c_next_tls_key"), int_lit(1)),
        ),
        int_lit(0),
    ]))
}

pub fn key_create_with_destructor(key: Expression, destructor: Expression) -> Expression {
    let target = arg_target(key);
    expr(ExprKind::Sequence(vec![
        assign_expr(target.clone(), ident("__c_next_tls_key")),
        assign_expr(ident("__c_last_tls_key"), target.clone()),
        assign_expr(index_expr(ident("__c_tls_destructors"), target), destructor),
        assign_expr(
            ident("__c_next_tls_key"),
            add(ident("__c_next_tls_key"), int_lit(1)),
        ),
        int_lit(0),
    ]))
}

pub fn set_specific(key: Expression, value: Expression) -> Expression {
    expr(ExprKind::Sequence(vec![
        assign_expr(index_expr(ident("__c_tls_values"), key), value.clone()),
        assign_expr(
            index_expr(ident("__c_tls_values"), ident("__c_last_tls_key")),
            value,
        ),
        int_lit(0),
    ]))
}

pub fn get_specific(key: Expression) -> Expression {
    expr(ExprKind::NullCoalesce {
        left: Box::new(index_expr(ident("__c_tls_values"), key)),
        right: Box::new(null_lit()),
    })
}

fn sem_target(sem: Expression) -> Expression {
    if is_direct_sem_arg(&sem) {
        return arg_target(sem);
    }
    index_expr(ident("__c_sem_values"), arg_target(sem))
}

/// The prelude half of the thread model — the three functions a spawned
/// thread and its joiner run. Emitted once by `c_runtime::prelude`.
pub fn runtime_functions() -> Vec<Statement> {
    vec![thread_entry_fn(), thread_join_fn(), barrier_wait_fn()]
}

/// `__c_thread_entry(h)` — every spawned thread starts HERE, not at the C
/// start function. It is what the shared spawn lowering can carry: one
/// funcref plus one i32 (the handle). It unpacks the real start function
/// and its `void*` from the shared tables, publishes the return value where
/// the joiner will look, and runs the TLS destructors that end a thread.
///
/// The fresh `__c_tls_values` is the whole of pthread TLS isolation: the
/// child VM holds a CLONE of the globals, so rebinding the slot here is
/// private to this thread while the parent keeps its own.
fn thread_entry_fn() -> Statement {
    let call_start = expr(ExprKind::Call {
        callee: Box::new(index_expr(ident("__c_thread_starts"), ident("h"))),
        args: vec![vybe_ast::Argument::positional(index_expr(
            ident("__c_thread_args"),
            ident("h"),
        ))],
        optional: false,
    });
    function_stmt(
        "__c_thread_entry",
        vec!["h"],
        vec![
            stmt(StmtKind::Expr(assign_expr(
                ident("__c_tls_values"),
                empty_object(),
            ))),
            stmt(StmtKind::Expr(assign_expr(
                index_expr(ident("__c_thread_results"), ident("h")),
                call_start,
            ))),
            stmt(StmtKind::Expr(tls_destructor_pass())),
            stmt(StmtKind::Expr(tls_destructor_pass())),
            stmt(StmtKind::Expr(tls_destructor_pass())),
            stmt(StmtKind::Expr(tls_destructor_pass())),
            stmt(StmtKind::Return(Some(int_lit(0)))),
        ],
    )
}

/// `__c_thread_join_h(h)` — wait for the task, then answer the start
/// function's return value. A cancelled thread is never waited on: we
/// cannot stop a running thread from outside, so `pthread_cancel` records
/// PTHREAD_CANCELED (-1) and the joiner takes it without blocking.
fn thread_join_fn() -> Statement {
    function_stmt(
        "__c_thread_join_h",
        vec!["h"],
        vec![
            if_stmt(
                eq(
                    index_expr(ident("__c_thread_results"), ident("h")),
                    int_lit(-1),
                ),
                vec![stmt(StmtKind::Return(Some(int_lit(-1))))],
                None,
            ),
            var_decl_stmt("task", index_expr(ident("__c_thread_tasks"), ident("h"))),
            if_stmt(
                binary(BinOp::NotEq, ident("task"), null_lit()),
                vec![stmt(StmtKind::Expr(expr(ExprKind::Call {
                    callee: Box::new(ident("__c_task_wait")),
                    args: vec![vybe_ast::Argument::positional(ident("task"))],
                    optional: false,
                })))],
                None,
            ),
            stmt(StmtKind::Return(Some(index_expr(
                ident("__c_thread_results"),
                ident("h"),
            )))),
        ],
    )
}

/// `__c_barrier_wait_h(h)` — arrive at the barrier, then block until the
/// generation turns over. The last arrival resets the tally, bumps the
/// generation and is the one thread that gets PTHREAD_BARRIER_SERIAL_THREAD.
/// Waiters yield through a real WASI sleep rather than burning a core.
fn barrier_wait_fn() -> Statement {
    function_stmt(
        "__c_barrier_wait_h",
        vec!["h"],
        vec![
            var_decl_stmt("limit", index_expr(ident("__c_barrier_limit"), ident("h"))),
            if_stmt(
                eq(ident("limit"), null_lit()),
                vec![stmt(StmtKind::Return(Some(int_lit(0))))],
                None,
            ),
            var_decl_stmt("gen", index_expr(ident("__c_barrier_gen"), ident("h"))),
            var_decl_stmt(
                "n",
                add(
                    expr(ExprKind::NullCoalesce {
                        left: Box::new(index_expr(ident("__c_barrier_arrived"), ident("h"))),
                        right: Box::new(int_lit(0)),
                    }),
                    int_lit(1),
                ),
            ),
            stmt(StmtKind::Expr(assign_expr(
                index_expr(ident("__c_barrier_arrived"), ident("h")),
                ident("n"),
            ))),
            if_stmt(
                binary(BinOp::GtEq, ident("n"), ident("limit")),
                vec![
                    stmt(StmtKind::Expr(assign_expr(
                        index_expr(ident("__c_barrier_arrived"), ident("h")),
                        int_lit(0),
                    ))),
                    stmt(StmtKind::Expr(assign_expr(
                        index_expr(ident("__c_barrier_gen"), ident("h")),
                        add(ident("gen"), int_lit(1)),
                    ))),
                    stmt(StmtKind::Return(Some(int_lit(1)))),
                ],
                None,
            ),
            var_decl_stmt("spins", int_lit(0)),
            stmt(StmtKind::While {
                cond: eq(
                    index_expr(ident("__c_barrier_gen"), ident("h")),
                    ident("gen"),
                ),
                body: vec![
                    stmt(StmtKind::Expr(expr(ExprKind::Call {
                        callee: Box::new(ident("__c_sleep_ms")),
                        args: vec![vybe_ast::Argument::positional(int_lit(1))],
                        optional: false,
                    }))),
                    stmt(StmtKind::Expr(assign_expr(
                        ident("spins"),
                        add(ident("spins"), int_lit(1)),
                    ))),
                    // A peer that never arrives is a deadlocked barrier, not
                    // a reason to hang the process forever.
                    if_stmt(
                        binary(BinOp::Gt, ident("spins"), int_lit(5000)),
                        vec![stmt(StmtKind::Return(Some(int_lit(0))))],
                        None,
                    ),
                ],
                else_body: None,
            }),
            stmt(StmtKind::Return(Some(int_lit(0)))),
        ],
    )
}

fn tls_destructor_pass() -> Expression {
    let key = ident("__c_last_tls_key");
    let value = index_expr(ident("__c_tls_values"), key.clone());
    let destructor = index_expr(ident("__c_tls_destructors"), key);
    expr(ExprKind::Ternary {
        cond: Box::new(and(
            binary(BinOp::NotEq, destructor.clone(), null_lit()),
            binary(BinOp::NotEq, value.clone(), null_lit()),
        )),
        then: Box::new(expr(ExprKind::Sequence(vec![
            assign_expr(
                index_expr(ident("__c_tls_values"), ident("__c_last_tls_key")),
                null_lit(),
            ),
            expr(ExprKind::Call {
                callee: Box::new(destructor),
                args: vec![vybe_ast::Argument::positional(value)],
                optional: false,
            }),
            int_lit(0),
        ]))),
        else_: Box::new(int_lit(0)),
    })
}

fn empty_object() -> Expression {
    expr(ExprKind::Object(vec![]))
}

fn sem_value(sem: Expression) -> Expression {
    if is_direct_sem_arg(&sem) {
        return arg_target(sem);
    }
    expr(ExprKind::NullCoalesce {
        left: Box::new(sem_target(sem)),
        right: Box::new(int_lit(0)),
    })
}

fn is_direct_sem_arg(value: &Expression) -> bool {
    match &value.kind {
        ExprKind::Unary {
            op: UnaryOp::AddrOf,
            ..
        } => true,
        ExprKind::RefOf(_) => true,
        ExprKind::Cast { expr, .. } => is_direct_sem_arg(expr),
        _ => false,
    }
}

fn attr_target(attr: Expression) -> Expression {
    expr(ExprKind::NullCoalesce {
        left: Box::new(arg_target(attr)),
        right: Box::new(int_lit(0)),
    })
}

fn arg_target(value: Expression) -> Expression {
    match value.kind {
        ExprKind::Cast { expr, .. } => arg_target(*expr),
        ExprKind::RefLoad(expr) => arg_target(*expr),
        ExprKind::Unary {
            op: UnaryOp::AddrOf,
            expr,
        } => *expr,
        ExprKind::RefOf(place) => match *place {
            PlaceExpr::Ident(name) => ident(&name),
            PlaceExpr::Member {
                object,
                field,
                null_safe,
            } => expr(ExprKind::Member {
                object,
                field,
                null_safe,
            }),
            PlaceExpr::Index {
                object,
                index,
                null_safe,
            } => expr(ExprKind::Index {
                object,
                index,
                null_safe,
            }),
            PlaceExpr::Deref(inner) => *inner,
        },
        _ => value,
    }
}

fn eq(left: Expression, right: Expression) -> Expression {
    binary(BinOp::Eq, left, right)
}

fn lt(left: Expression, right: Expression) -> Expression {
    binary(BinOp::Lt, left, right)
}

fn gt(left: Expression, right: Expression) -> Expression {
    binary(BinOp::Gt, left, right)
}

fn add(left: Expression, right: Expression) -> Expression {
    binary(BinOp::Add, left, right)
}

fn and(left: Expression, right: Expression) -> Expression {
    binary(BinOp::And, left, right)
}

fn or(left: Expression, right: Expression) -> Expression {
    binary(BinOp::Or, left, right)
}

fn not(value: Expression) -> Expression {
    expr(ExprKind::Unary {
        op: UnaryOp::Not,
        expr: Box::new(value),
    })
}

fn bit_set(value: Expression, mask: i64) -> Expression {
    binary(
        BinOp::NotEq,
        binary(BinOp::BitAnd, value, int_lit(mask)),
        int_lit(0),
    )
}

fn binary(op: BinOp, left: Expression, right: Expression) -> Expression {
    expr(ExprKind::Binary {
        op,
        left: Box::new(left),
        right: Box::new(right),
    })
}

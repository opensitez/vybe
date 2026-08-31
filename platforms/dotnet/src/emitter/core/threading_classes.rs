//! `System.Threading` coordination primitives, synthesized as REAL classes.
//!
//! Same move, same reason as [`super::interop_classes`]: a `ClassType` in
//! `class_exports()` cannot be constructed with `new` from user code, and the
//! failure it produces is always `undefined is not callable`. Injecting a
//! `StmtKind::ClassDecl` gives these types a real constructor, real fields and
//! ordinary dispatch, so `using var cde = new CountdownEvent(2)` binds its
//! `Dispose` the same way any user class does.
//!
//! ⚠ **These are SINGLE-THREADED-DETERMINISTIC on purpose.** `SpinWait`,
//! `CountdownEvent` and friends coordinate threads in .NET, but every corpus
//! use is a counter observed from one thread — `SpinOnce()` then `Count`,
//! `Signal()` twice then `Wait(100)`. Modelling the counter exactly is what
//! those programs actually observe, and it is honest: a `Wait` that blocks for
//! a signal which can never arrive on a single thread would hang rather than
//! answer. Where a real scheduler becomes observable this must be revisited,
//! and the tests that would notice are the ones that spawn threads.
//!
//! ⚠ Injection is GATED on the program naming the type, for the reason
//! `interop_classes` documents: unconditional injection shifts typeidx
//! numbering for every language while the class model is mid-conversion.

use super::interop_classes::{
    assign, by_ref_param, class, getter, me, method, param, typed_field,
};
use vybe_ast::{
    BinOp, ClassMember, ConstructorInitializerTarget, ExprKind, Expression, Literal, Param, Span,
    Statement, StmtKind, Visibility,
};

/// Whether `name` is a class this module injects.
pub fn is_synthesized_threading_class(name: &str) -> bool {
    matches!(
        name.to_ascii_lowercase().as_str(),
        "spinwait"
            | "countdownevent"
            | "manualreseteventslim"
            | "semaphoreslim"
            | "spinlock"
            | "barrier"
            | "readerwriterlockslim"
            | "periodictimer"
            | "threadlocal"
            | "lock"

    )
}

/// The threading classes, or nothing when `source` never names one.
pub fn synthesize_threading_classes(source: &str) -> Vec<Statement> {
    let lowered = source.to_ascii_lowercase();
    let mut out = Vec::new();
    if lowered.contains("spinwait") {
        out.push(spin_wait_class());
    }
    if lowered.contains("countdownevent") {
        out.push(countdown_event_class());
    }
    if lowered.contains("manualreseteventslim") {
        out.push(manual_reset_event_slim_class());
    }
    if lowered.contains("semaphoreslim") {
        out.push(semaphore_full_exception_class());
        out.push(semaphore_slim_class());
    }
    if lowered.contains("spinlock") {
        out.push(spin_lock_class());
    }
    // ⛔ GATED ON THE QUALIFIED SPELLING, NOT ON `"lock"`. Every other type here
    // has a name distinctive enough to match as a bare substring; `lock` is a
    // C# KEYWORD and a fragment of `unlock`, `block`, `deadlock` and `locked`,
    // so `lowered.contains("lock")` would inject this class into a large part
    // of the corpus. `System.Threading.Lock` (C# 13) is how the type is
    // written, and `new Lock(` covers the `using System.Threading;` form.
    if lowered.contains("threading.lock") || lowered.contains("new lock(") {
        out.push(lock_class());
    }
    if lowered.contains("barrier") {
        out.push(barrier_class());
    }
    if lowered.contains("readerwriterlockslim") {
        out.push(reader_writer_lock_slim_class());
    }
    if lowered.contains("periodictimer") {
        out.push(periodic_timer_class());
    }
    if lowered.contains("threadlocal") {
        out.push(thread_local_class());
    }
    // ⛔ NO `ManualResetEvent` / `WaitHandle` HERE. Both are already
    // `ClassType` exports in `component_classes_threading.rs`, and a second
    // home does not win — it loses silently: `--dump-classes` showed my
    // synthesized `WaitHandle` with `members: (none)`, the export having
    // shadowed it, while `ManualResetEvent(true).WaitOne()` answered False
    // because the EXPORT's constructor ran, not mine. Fix the registration
    // that exists; never add a rival to it.
    out
}

// ── local AST helpers ───────────────────────────────────────────────────

fn int_lit(v: i64) -> Expression {
    Expression::new(ExprKind::Lit(Literal::Int(v)))
}

fn bool_lit(v: bool) -> Expression {
    Expression::new(ExprKind::Lit(Literal::Bool(v)))
}

fn ident(name: &str) -> Expression {
    Expression::new(ExprKind::Ident(name.into()))
}

fn binary(left: Expression, op: BinOp, right: Expression) -> Expression {
    Expression::new(ExprKind::Binary {
        op,
        left: Box::new(left),
        right: Box::new(right),
    })
}

/// `Throw New <exception>(message)`.
///
/// ⛔ .NET THROWS ON THESE PATHS AND SILENCE IS THE WRONG ANSWER. A differential
/// against the real SDK found both: `CountdownEvent.Signal()` past zero and
/// `SemaphoreSlim.Release()` past `maxCount` are `InvalidOperationException` and
/// `SemaphoreFullException` respectively, and the first cut of this file
/// returned a cheerful `True` for one and silently over-counted for the other.
/// NO corpus test covers either, which is exactly why they had to be checked
/// against dotnet rather than against the suite.
fn throw_new(exception: &str, message: &str) -> Statement {
    Statement::new(StmtKind::Throw {
        expr: Some(Expression::new(ExprKind::New {
            class: Box::new(ident(exception)),
            args: vec![vybe_ast::Argument::positional(Expression::new(
                ExprKind::Lit(Literal::Str(message.into())),
            ))],
        })),
        cause: None,
    })
}

/// `If <cond> Then <body>`.
fn if_then(cond: Expression, body: Vec<Statement>) -> Statement {
    Statement::with_span(
        StmtKind::If {
            cond,
            then_body: body,
            elifs: Vec::new(),
            else_body: None,
        },
        Span::default(),
    )
}

fn ret(value: Expression) -> Statement {
    Statement::new(StmtKind::Return(Some(value)))
}

/// A constructor taking `params`, whose body is `body`.
fn ctor(params: Vec<Param>, body: Vec<Statement>) -> ClassMember {
    ClassMember::Constructor {
        name: None,
        params,
        body,
        base_args: None,
        initializer_target: ConstructorInitializerTarget::Base,
        visibility: Visibility::Public,
    }
}

/// `Dispose()` — present so `using var x = new T(...)` binds. Every one of
/// these types is `IDisposable` in .NET and none of them holds an OS handle
/// here, so the body is deliberately empty rather than absent: a MISSING
/// `Dispose` makes the `using` declaration itself fail, which reads as the
/// constructor being broken.
fn dispose() -> ClassMember {
    method("Dispose", Vec::new(), Vec::new(), true)
}

// ── The classes ─────────────────────────────────────────────────────────

/// `System.Threading.SpinWait` — a spin counter.
///
/// `SpinOnce()` increments `Count`; .NET yields the timeslice, which is not
/// observable here. `NextSpinWillYield` is true once the count passes the
/// threshold .NET uses before it starts yielding (20).
fn spin_wait_class() -> Statement {
    let count = "Count";
    class(
        "SpinWait",
        Vec::new(),
        vec![
            typed_field(count, "Integer"),
            ctor(Vec::new(), vec![assign(count, int_lit(0))]),
            method(
                "SpinOnce",
                Vec::new(),
                vec![assign(
                    count,
                    binary(me(count), BinOp::Add, int_lit(1)),
                )],
                true,
            ),
            method("Reset", Vec::new(), vec![assign(count, int_lit(0))], true),
            getter(
                "NextSpinWillYield",
                vec![ret(binary(me(count), BinOp::GtEq, int_lit(20)))],
            ),
        ],
    )
}

/// `System.Threading.CountdownEvent` — a counter that reaches zero.
///
/// `Wait(timeout)` answers whether the count has already reached zero rather
/// than blocking: on a single thread a signal that has not arrived cannot
/// arrive during the wait, so answering immediately is both correct for every
/// program that can observe it and strictly better than hanging.
fn countdown_event_class() -> Statement {
    let current = "CurrentCount";
    let initial = "InitialCount";
    class(
        "CountdownEvent",
        Vec::new(),
        vec![
            typed_field(current, "Integer"),
            typed_field(initial, "Integer"),
            ctor(
                vec![param("initialCount", int_lit(0))],
                vec![
                    assign(current, ident("initialCount")),
                    assign(initial, ident("initialCount")),
                ],
            ),
            // `Signal()` returns True when the count reached zero on THIS call,
            // and THROWS when it is already zero — verified against the SDK.
            method(
                "Signal",
                Vec::new(),
                vec![
                    if_then(
                        binary(me(current), BinOp::LtEq, int_lit(0)),
                        vec![throw_new(
                            "InvalidOperationException",
                            "Invalid attempt made to decrement the event's count below zero.",
                        )],
                    ),
                    assign(current, binary(me(current), BinOp::Sub, int_lit(1))),
                    ret(binary(me(current), BinOp::LtEq, int_lit(0))),
                ],
                false,
            ),
            method(
                "AddCount",
                vec![param("count", int_lit(1))],
                vec![assign(
                    current,
                    binary(me(current), BinOp::Add, ident("count")),
                )],
                true,
            ),
            method(
                "Reset",
                Vec::new(),
                vec![assign(current, me(initial))],
                true,
            ),
            method(
                "Wait",
                vec![param("timeout", int_lit(0))],
                vec![ret(binary(me(current), BinOp::LtEq, int_lit(0)))],
                false,
            ),
            getter(
                "IsSet",
                vec![ret(binary(me(current), BinOp::LtEq, int_lit(0)))],
            ),
            dispose(),
        ],
    )
}

/// `System.Threading.ManualResetEventSlim` — a latch.
fn manual_reset_event_slim_class() -> Statement {
    let set = "IsSet";
    class(
        "ManualResetEventSlim",
        Vec::new(),
        vec![
            typed_field(set, "Boolean"),
            ctor(
                vec![param("initialState", bool_lit(false))],
                vec![assign(set, ident("initialState"))],
            ),
            method("Set", Vec::new(), vec![assign(set, bool_lit(true))], true),
            method("Reset", Vec::new(), vec![assign(set, bool_lit(false))], true),
            method(
                "Wait",
                vec![param("timeout", int_lit(0))],
                vec![ret(me(set))],
                false,
            ),
            dispose(),
        ],
    )
}

/// `System.Threading.SemaphoreSlim` — a permit count.
fn semaphore_slim_class() -> Statement {
    let count = "CurrentCount";
    class(
        "SemaphoreSlim",
        Vec::new(),
        vec![
            typed_field(count, "Integer"),
            typed_field("__max_count", "Integer"),
            ctor(
                vec![
                    param("initialCount", int_lit(0)),
                    // .NET's one-arg form is unbounded; `int.MaxValue` is the
                    // documented default and makes the bound check below inert
                    // for it rather than needing a second code path.
                    param("maxCount", int_lit(2147483647)),
                ],
                vec![
                    assign(count, ident("initialCount")),
                    assign("__max_count", ident("maxCount")),
                ],
            ),
            // `Wait` takes a permit when one is available and reports whether
            // it got one; the timeout overload returns that Boolean, the plain
            // one discards it, and both run the same body.
            method(
                "Wait",
                vec![param("timeout", int_lit(0))],
                vec![
                    Statement::with_span(
                        StmtKind::If {
                            cond: binary(me(count), BinOp::Gt, int_lit(0)),
                            then_body: vec![
                                assign(count, binary(me(count), BinOp::Sub, int_lit(1))),
                                ret(bool_lit(true)),
                            ],
                            elifs: Vec::new(),
                            else_body: None,
                        },
                        Span::default(),
                    ),
                    ret(bool_lit(false)),
                ],
                false,
            ),
            method(
                "Release",
                vec![param("releaseCount", int_lit(1))],
                vec![
                    if_then(
                        binary(
                            binary(me(count), BinOp::Add, ident("releaseCount")),
                            BinOp::Gt,
                            me("__max_count"),
                        ),
                        vec![throw_new(
                            "SemaphoreFullException",
                            "Adding the specified count to the semaphore would cause it to exceed its maximum count.",
                        )],
                    ),
                    assign(count, binary(me(count), BinOp::Add, ident("releaseCount"))),
                ],
                true,
            ),
            dispose(),
        ],
    )
}

/// `System.Threading.SemaphoreFullException` — not in the shared exception
/// hierarchy, so it is injected here alongside its only thrower.
fn semaphore_full_exception_class() -> Statement {
    class("SemaphoreFullException", vec!["Exception".into()], Vec::new())
}

/// `System.Threading.SpinLock` — an uncontended lock flag.
///
/// ⚠ `Enter` reports through a ByRef `lockTaken` in .NET. That parameter is
/// declared here so the call binds; the write reaching the caller's variable is
/// what `by_ref_param` exists for in `interop_classes`, and is wired the same
/// way when a corpus program observes it.
fn spin_lock_class() -> Statement {
    let held = "__held";
    class(
        "SpinLock",
        Vec::new(),
        vec![
            typed_field(held, "Boolean"),
            ctor(
                vec![param("enableThreadOwnerTracking", bool_lit(true))],
                vec![assign(held, bool_lit(false))],
            ),
            // ⛔ `lockTaken` IS ByRef AND MUST BE WRITTEN. .NET reports the
            // acquisition through the argument, and the corpus reads the
            // CALLER's variable straight after. A by-value param compiles and
            // runs, leaves `lockTaken` false, and the failure surfaces as the
            // test's own assertion throwing — not as anything naming `Enter`.
            method(
                "Enter",
                vec![by_ref_param("lockTaken")],
                vec![
                    assign(held, bool_lit(true)),
                    Statement::new(StmtKind::Assign {
                        targets: vec![ident("lockTaken")],
                        value: bool_lit(true),
                        by_ref: false,
                    }),
                ],
                true,
            ),
            method(
                "TryEnter",
                vec![by_ref_param("lockTaken")],
                vec![
                    assign(held, bool_lit(true)),
                    Statement::new(StmtKind::Assign {
                        targets: vec![ident("lockTaken")],
                        value: bool_lit(true),
                        by_ref: false,
                    }),
                ],
                true,
            ),
            method("Exit", Vec::new(), vec![assign(held, bool_lit(false))], true),
            getter("IsHeld", vec![ret(me(held))]),
        ],
    )
}

/// `System.Threading.Barrier` — a phase counter.
///
/// `SignalAndWait` completes a phase when every participant has signalled. With
/// the single participant every corpus program uses, each call completes one
/// phase, so `CurrentPhaseNumber` counts calls. The participant arithmetic is
/// modelled rather than assumed so a two-participant program answers sensibly
/// instead of silently counting wrong.
fn barrier_class() -> Statement {
    let phase = "CurrentPhaseNumber";
    let participants = "ParticipantCount";
    let signalled = "__signalled";
    class(
        "Barrier",
        Vec::new(),
        vec![
            typed_field(phase, "Integer"),
            typed_field(participants, "Integer"),
            typed_field(signalled, "Integer"),
            ctor(
                vec![param("participantCount", int_lit(1))],
                vec![
                    assign(phase, int_lit(0)),
                    assign(participants, ident("participantCount")),
                    assign(signalled, int_lit(0)),
                ],
            ),
            method(
                "SignalAndWait",
                vec![param("timeout", int_lit(0))],
                vec![
                    assign(signalled, binary(me(signalled), BinOp::Add, int_lit(1))),
                    if_then(
                        binary(me(signalled), BinOp::GtEq, me(participants)),
                        vec![
                            assign(signalled, int_lit(0)),
                            assign(phase, binary(me(phase), BinOp::Add, int_lit(1))),
                        ],
                    ),
                    ret(bool_lit(true)),
                ],
                false,
            ),
            method(
                "AddParticipant",
                Vec::new(),
                vec![assign(
                    participants,
                    binary(me(participants), BinOp::Add, int_lit(1)),
                )],
                true,
            ),
            method(
                "RemoveParticipant",
                Vec::new(),
                vec![assign(
                    participants,
                    binary(me(participants), BinOp::Sub, int_lit(1)),
                )],
                true,
            ),
            dispose(),
        ],
    )
}

/// `System.Threading.ReaderWriterLockSlim` — held-flags.
fn reader_writer_lock_slim_class() -> Statement {
    let read = "IsReadLockHeld";
    let write = "IsWriteLockHeld";
    let upgradeable = "IsUpgradeableReadLockHeld";
    class(
        "ReaderWriterLockSlim",
        Vec::new(),
        vec![
            typed_field(read, "Boolean"),
            typed_field(write, "Boolean"),
            typed_field(upgradeable, "Boolean"),
            ctor(
                vec![param("recursionPolicy", int_lit(0))],
                vec![
                    assign(read, bool_lit(false)),
                    assign(write, bool_lit(false)),
                    assign(upgradeable, bool_lit(false)),
                ],
            ),
            method("EnterReadLock", Vec::new(), vec![assign(read, bool_lit(true))], true),
            method("ExitReadLock", Vec::new(), vec![assign(read, bool_lit(false))], true),
            method("EnterWriteLock", Vec::new(), vec![assign(write, bool_lit(true))], true),
            method("ExitWriteLock", Vec::new(), vec![assign(write, bool_lit(false))], true),
            method(
                "EnterUpgradeableReadLock",
                Vec::new(),
                vec![assign(upgradeable, bool_lit(true))],
                true,
            ),
            method(
                "ExitUpgradeableReadLock",
                Vec::new(),
                vec![assign(upgradeable, bool_lit(false))],
                true,
            ),
            method(
                "TryEnterReadLock",
                vec![param("timeout", int_lit(0))],
                vec![assign(read, bool_lit(true)), ret(bool_lit(true))],
                false,
            ),
            method(
                "TryEnterWriteLock",
                vec![param("timeout", int_lit(0))],
                vec![assign(write, bool_lit(true)), ret(bool_lit(true))],
                false,
            ),
            dispose(),
        ],
    )
}

/// `System.Threading.PeriodicTimer` — constructible and disposable.
///
/// ⚠ `WaitForNextTickAsync` is deliberately NOT modelled: it is a real await on
/// a real timer, and answering it synchronously would be a lie a corpus program
/// could observe. Construction and disposal are what the tests exercise.
fn periodic_timer_class() -> Statement {
    class(
        "PeriodicTimer",
        Vec::new(),
        vec![
            typed_field("__period", "Object"),
            ctor(
                vec![param("period", int_lit(0))],
                vec![assign("__period", ident("period"))],
            ),
            dispose(),
        ],
    )
}

/// `System.Threading.ThreadLocal<T>` — a lazily-initialised per-thread slot.
///
/// One thread here, so the slot is the value. `Value` invokes the factory on
/// first read and caches, which is what `IsValueCreated` reports.
fn thread_local_class() -> Statement {
    let value = "__value";
    let created = "IsValueCreated";
    let factory = "__factory";
    class(
        "ThreadLocal",
        Vec::new(),
        vec![
            typed_field(value, "Object"),
            typed_field(created, "Boolean"),
            typed_field(factory, "Object"),
            ctor(
                vec![param("valueFactory", Expression::new(ExprKind::Lit(Literal::Null)))],
                vec![
                    assign(factory, ident("valueFactory")),
                    assign(created, bool_lit(false)),
                ],
            ),
            ClassMember::Property {
                name: "Value".into(),
                type_hint: None,
                getter: Some(vec![
                    if_then(
                        Expression::new(ExprKind::Unary {
                            op: vybe_ast::UnaryOp::Not,
                            expr: Box::new(me(created)),
                        }),
                        vec![
                            assign(
                                value,
                                Expression::new(ExprKind::Call {
                                    callee: Box::new(me(factory)),
                                    args: Vec::new(),
                                    optional: false,
                                }),
                            ),
                            assign(created, bool_lit(true)),
                        ],
                    ),
                    ret(me(value)),
                ]),
                setter: None,
                is_auto: false,
                modifiers: vybe_ast::Modifiers::default(),
            },
            dispose(),
        ],
    )
}

/// `System.Threading.Lock` — C# 13's dedicated lock object.
///
/// `lock (obj) { … }` already works on ANY object in this runtime, so what was
/// missing was only the TYPE: `new System.Threading.Lock()` answered
/// `undefined is not callable`. The members below are .NET 9's real surface.
///
/// ⚠ Single-threaded-deterministic, for the reason the module header gives:
/// `Enter`/`Exit` maintain the recursion count that `IsHeldByCurrentThread`
/// reports, and `TryEnter` always succeeds because on one thread it always
/// would. A contended `TryEnter` returning false is not expressible here.
fn lock_class() -> Statement {
    let depth = "__depth";
    class(
        "Lock",
        Vec::new(),
        vec![
            typed_field(depth, "Integer"),
            ctor(Vec::new(), vec![assign(depth, int_lit(0))]),
            method(
                "Enter",
                Vec::new(),
                vec![assign(
                    depth,
                    Expression::new(ExprKind::Binary {
                        op: BinOp::Add,
                        left: Box::new(me(depth)),
                        right: Box::new(int_lit(1)),
                    }),
                )],
                true,
            ),
            method(
                "Exit",
                Vec::new(),
                vec![assign(
                    depth,
                    Expression::new(ExprKind::Binary {
                        op: BinOp::Sub,
                        left: Box::new(me(depth)),
                        right: Box::new(int_lit(1)),
                    }),
                )],
                true,
            ),
            method(
                "TryEnter",
                Vec::new(),
                vec![
                    assign(
                        depth,
                        Expression::new(ExprKind::Binary {
                            op: BinOp::Add,
                            left: Box::new(me(depth)),
                            right: Box::new(int_lit(1)),
                        }),
                    ),
                    Statement::new(StmtKind::Return(Some(bool_lit(true)))),
                ],
                false,
            ),
            // `using (lk.EnterScope())` — the scope IS the lock here, and its
            // `Dispose` releases, which is the observable contract.
            method(
                "EnterScope",
                Vec::new(),
                vec![
                    assign(
                        depth,
                        Expression::new(ExprKind::Binary {
                            op: BinOp::Add,
                            left: Box::new(me(depth)),
                            right: Box::new(int_lit(1)),
                        }),
                    ),
                    Statement::new(StmtKind::Return(Some(Expression::new(ExprKind::This)))),
                ],
                false,
            ),
            method(
                "Dispose",
                Vec::new(),
                vec![assign(
                    depth,
                    Expression::new(ExprKind::Binary {
                        op: BinOp::Sub,
                        left: Box::new(me(depth)),
                        right: Box::new(int_lit(1)),
                    }),
                )],
                true,
            ),
            getter(
                "IsHeldByCurrentThread",
                vec![Statement::new(StmtKind::Return(Some(Expression::new(
                    ExprKind::Binary {
                        op: BinOp::Gt,
                        left: Box::new(me(depth)),
                        right: Box::new(int_lit(0)),
                    },
                ))))],
            ),
        ],
    )
}

//! `threading` — the synchronisation primitives.
//!
//! These are the *shape* of the primitives, not working concurrency: there is
//! one thread of execution, so a lock always acquires, a `Condition.wait`
//! returns immediately and a `Barrier` never blocks. That was true of the
//! prelude too — what changes is that they are declared classes, so
//! `with lock:` binds through the real `__enter__`/`__exit__` slots and
//! `isinstance(l, Lock)` answers from the rtt.
//!
//! ⛔ The THREAD classes are deliberately NOT here: `Thread`/`Timer` and the
//! executor family depend on `__py_thread_*` helpers whose semantics are the
//! open item in [[project_python_threads_never_reach_the_event_loop]] (bodies
//! run at `join()`). Converting the primitives is separable from fixing that,
//! and mixing the two would make a runtime bug look like a conversion
//! regression.

use super::builders::*;
use vybe_ast::{BinOp, Statement};

/// `acquire` / `release` / `__enter__` / `__exit__` — the context-manager pair
/// every primitive here shares. `__enter__` answers `self` so
/// `with lock as l:` binds.
fn enter_exit() -> Vec<vybe_ast::ClassMember> {
    vec![
        method(
            "__enter__",
            vec![],
            vec![
                expr_stmt(call(member(ident("self"), "acquire"), vec![])),
                ret(ident("self")),
            ],
        ),
        method(
            "__exit__",
            any_args(),
            vec![
                expr_stmt(call(member(ident("self"), "release"), vec![])),
                ret(bool_lit(false)),
            ],
        ),
    ]
}

pub(super) fn base_lock() -> Statement {
    let mut members = vec![
        init(vec![], vec![set_this("locked", bool_lit(false))]),
        // ⛔⛔ THE SHARED SPINLOCK DEADLOCKS HERE — MEASURED, DO NOT RETRY
        // WITHOUT FIXING WAIT/NOTIFY FIRST.
        //
        // `primitives::threading::emit_lock_acquire` is the right code and is
        // what C# `lock {}` / VB `SyncLock` compile to: `atomic_rmw_xchg` to
        // take, then `memory.atomic.wait32` when the word is held. It is wired
        // and ready in `emitter/lock_adapter.rs` behind `__py_lock_acquire` /
        // `__py_lock_release`.
        //
        // Four threads incrementing under `with lock:` HANG in `join`: a
        // thread that parks in `wait32` is never notified, so the holder never
        // runs to release. Uncontended `with lock:` on the main thread works,
        // which is exactly why this looks fine until a second thread arrives.
        //
        // So the flag stays for now. It is not a lock — it records state so
        // `locked` reads correctly — and it cannot deadlock. The real spinlock
        // goes in the moment `wait32`/`notify` wake a parked python thread.
        method(
            "acquire",
            vec![
                param("blocking", Some(bool_lit(true))),
                param("timeout", Some(num(-1.0))),
            ],
            vec![set_this("locked", bool_lit(true)), ret(bool_lit(true))],
        ),
        method(
            "release",
            vec![],
            vec![set_this("locked", bool_lit(false))],
        ),
    ];
    members.extend(enter_exit());
    class("__PyLock", members)
}

/// `Lock` and `RLock` — `__PyLock` with nothing added. The parent IS the
/// declaration.
pub(super) const LOCK_ALIASES: &[(&str, &str)] = &[("RLock", "__PyLock"), ("Lock", "__PyLock")];

pub(super) fn lock_alias(name: &'static str, parent: &'static str) -> Statement {
    class_extending(name, &[parent], vec![])
}

pub(super) fn semaphore() -> Statement {
    let mut members = vec![
        init(
            vec![param("value", Some(num(1.0)))],
            vec![
                set_this("_value", ident("value")),
                set_this("_initial_value", ident("value")),
            ],
        ),
        method(
            "acquire",
            vec![
                param("blocking", Some(bool_lit(true))),
                param("timeout", Some(null())),
            ],
            vec![
                if_stmt(
                    binary(BinOp::Gt, this_field("_value"), num(0.0)),
                    vec![
                        set_this(
                            "_value",
                            binary(BinOp::Sub, this_field("_value"), num(1.0)),
                        ),
                        ret(bool_lit(true)),
                    ],
                ),
                ret(bool_lit(false)),
            ],
        ),
        method(
            "release",
            vec![param("n", Some(num(1.0)))],
            vec![set_this("_value", add(this_field("_value"), ident("n")))],
        ),
    ];
    members.extend(enter_exit());
    class("Semaphore", members)
}

/// `BoundedSemaphore` overrides only `release`, to cap at the initial value —
/// which is the one behavioural difference CPython documents.
pub(super) fn bounded_semaphore() -> Statement {
    class_extending(
        "BoundedSemaphore",
        &["Semaphore"],
        vec![method(
            "release",
            vec![param("n", Some(num(1.0)))],
            vec![
                set_this("_value", add(this_field("_value"), ident("n"))),
                if_stmt(
                    binary(BinOp::Gt, this_field("_value"), this_field("_initial_value")),
                    vec![set_this("_value", this_field("_initial_value"))],
                ),
            ],
        )],
    )
}

pub(super) fn event() -> Statement {
    class(
        "Event",
        vec![
            init(vec![], vec![set_this("_flag", bool_lit(false))]),
            method("is_set", vec![], vec![ret(this_field("_flag"))]),
            method("set", vec![], vec![set_this("_flag", bool_lit(true))]),
            method("clear", vec![], vec![set_this("_flag", bool_lit(false))]),
            method(
                "wait",
                vec![param("timeout", Some(null()))],
                vec![ret(this_field("_flag"))],
            ),
        ],
    )
}

pub(super) fn condition() -> Statement {
    let mut members = vec![
        init(
            vec![param("lock", Some(null()))],
            vec![
                set_this("_lock", new("RLock", vec![])),
                if_stmt(
                    is_not_none(ident("lock")),
                    vec![set_this("_lock", ident("lock"))],
                ),
            ],
        ),
        method(
            "acquire",
            any_args(),
            vec![
                // ⛔ Bind the receiver: `self._lock.acquire()` is a call on a
                // nested expression, which is the shape that silently does
                // nothing.
                assign(ident("__l"), this_field("_lock")),
                ret(call(member(ident("__l"), "acquire"), vec![])),
            ],
        ),
        method(
            "release",
            vec![],
            vec![
                assign(ident("__l"), this_field("_lock")),
                ret(call(member(ident("__l"), "release"), vec![])),
            ],
        ),
        method(
            "wait",
            vec![param("timeout", Some(null()))],
            vec![ret(bool_lit(true))],
        ),
        method("notify", vec![param("n", Some(num(1.0)))], vec![ret(null())]),
        method("notify_all", vec![], vec![ret(null())]),
    ];
    members.extend(enter_exit());
    class("Condition", members)
}

pub(super) fn barrier() -> Statement {
    class(
        "Barrier",
        vec![
            init(
                vec![
                    param("parties", None),
                    param("action", Some(null())),
                    param("timeout", Some(null())),
                ],
                vec![
                    set_this("parties", ident("parties")),
                    set_this("n_waiting", num(0.0)),
                    set_this("broken", bool_lit(false)),
                ],
            ),
            method(
                "wait",
                vec![param("timeout", Some(null()))],
                vec![ret(num(0.0))],
            ),
            method("reset", vec![], vec![set_this("n_waiting", num(0.0))]),
            method("abort", vec![], vec![set_this("broken", bool_lit(true))]),
        ],
    )
}

/// `threading.local()` — a bare object whose attributes are per-thread. With
/// one thread it is just an object, which is exactly what the class says.
pub(super) fn thread_local() -> Statement {
    class("local", vec![init(vec![], vec![])])
}

/// `Thread`.
///
/// ⛔ The BODY is not a stub: `start`/`join` route to `__py_thread_start` /
/// `__py_thread_join`, which are profile builtins backed by the SHARED
/// `primitives::threading::emit_thread_spawn` / `emit_thread_join` —
/// real WASM-threads spawn and join. The deferral list around them
/// (`__py_pending_threads`) is the workaround for bodies not reaching the
/// event loop ([[project_python_threads_never_reach_the_event_loop]]); it is
/// carried across UNCHANGED so the conversion is separable from that fix.
pub(super) fn thread() -> Statement {
    class(
        "Thread",
        vec![
            init(
                vec![
                    param("group", Some(null())),
                    param("target", Some(null())),
                    param("name", Some(null())),
                    param("args", Some(list_of(vec![]))),
                    param("kwargs", Some(null())),
                    param("daemon", Some(null())),
                ],
                vec![
                    set_this("group", ident("group")),
                    set_this("_target", ident("target")),
                    set_this("name", str_lit("Thread")),
                    if_stmt(
                        is_not_none(ident("name")),
                        vec![set_this("name", ident("name"))],
                    ),
                    set_this("_args", ident("args")),
                    set_this("_kwargs", call_global("dict", vec![])),
                    if_stmt(
                        is_not_none(ident("kwargs")),
                        vec![set_this("_kwargs", ident("kwargs"))],
                    ),
                    set_this("daemon", is_true(ident("daemon"))),
                    set_this("_started", bool_lit(false)),
                    set_this("_done", bool_lit(false)),
                    set_this("_target_name", str_lit("")),
                ],
            ),
            method(
                "start",
                vec![],
                vec![expr_stmt(call_global("__py_thread_start", vec![ident("self")]))],
            ),
            method(
                "run",
                vec![],
                vec![expr_stmt(call_global("__py_thread_run", vec![ident("self")]))],
            ),
            method(
                "join",
                vec![param("timeout", Some(null()))],
                vec![expr_stmt(call_global(
                    "__py_thread_join",
                    vec![ident("self"), ident("timeout")],
                ))],
            ),
            method(
                "is_alive",
                vec![],
                vec![ret(binary(
                    BinOp::And,
                    this_field("_started"),
                    unary_not(this_field("_done")),
                ))],
            ),
        ],
    )
}

/// `Timer` — a `Thread` that remembers an interval. It does not schedule:
/// `Timer.start()` runs the target like any other thread.
pub(super) fn timer() -> Statement {
    class_extending(
        "Timer",
        &["Thread"],
        vec![init(
            vec![
                param("interval", None),
                param("function", None),
                param("args", Some(null())),
                param("kwargs", Some(null())),
            ],
            vec![
                set_this("interval", ident("interval")),
                set_this("_target", ident("function")),
                set_this("name", str_lit("Timer")),
                set_this("_args", list_of(vec![])),
                if_stmt(
                    is_not_none(ident("args")),
                    vec![set_this("_args", ident("args"))],
                ),
                set_this("_kwargs", call_global("dict", vec![])),
                set_this("daemon", bool_lit(false)),
                set_this("_started", bool_lit(false)),
                set_this("_done", bool_lit(false)),
                set_this("_target_name", str_lit("")),
            ],
        )],
    )
}

/// The module surface, plus the thread helpers `Thread` calls.
///
/// ⛔ `__py_thread_start_common` / `__py_thread_join_common` are PROFILE
/// BUILTINS backed by `primitives::threading::emit_thread_spawn` /
/// `emit_thread_join` — the real shared WASM-threads path. The pending list
/// around them exists because a spawned body does not reach the event loop
/// ([[project_python_threads_never_reach_the_event_loop]]), so `join` is where
/// the work actually runs. Carried across UNCHANGED: fixing that is the next
/// step, and mixing it into the conversion would make a runtime bug look like a
/// conversion regression.
pub(super) fn module_functions() -> Vec<Statement> {
    vec![
        global_assign("__py_pending_threads", list_of(vec![])),
        function(
            "__py_thread_call",
            vec![
                param("target", None),
                param("args", None),
                param("thread", Some(null())),
            ],
            vec![
                if_stmt(
                    is_none(ident("target")),
                    vec![ret(null())],
                ),
                expr_stmt(call_spread(ident("target"), ident("args"))),
                if_stmt(
                    is_not_none(ident("thread")),
                    vec![assign(
                        index(ident("thread"), str_lit("_done")),
                        bool_lit(true),
                    )],
                ),
                ret(null()),
            ],
        ),
        function(
            "__py_thread_run",
            vec![param("thread", None)],
            vec![
                if_stmt(
                    field_of(ident("thread"), "_done"),
                    vec![ret(null())],
                ),
                assign(index(ident("thread"), str_lit("_started")), bool_lit(true)),
                expr_stmt(call_global(
                    "__py_thread_call",
                    vec![
                        field_of(ident("thread"), "_target"),
                        field_of(ident("thread"), "_args"),
                        ident("thread"),
                    ],
                )),
                assign(index(ident("thread"), str_lit("_done")), bool_lit(true)),
            ],
        ),
        // The nullary closure the spawn is handed: it captures the thread and
        // forwards its own `_target` and `_args` when the worker runs it.
        function(
            "__py_thread_runner",
            vec![param("thread", None)],
            vec![
                function(
                    "__run",
                    vec![],
                    vec![expr_stmt(call_global(
                        "__py_thread_call",
                        vec![
                            field_of(ident("thread"), "_target"),
                            field_of(ident("thread"), "_args"),
                            ident("thread"),
                        ],
                    ))],
                ),
                ret(ident("__run")),
            ],
        ),
        // ▶▶ REAL SPAWN. `__py_thread_start_common` is a profile builtin
        // backed by `primitives::threading::emit_thread_spawn` — it grows the
        // funcref table, builds a worker chunk that invokes the stashed target,
        // and stores the task handle on the thread.
        //
        // ⛔ The target handed to it is the THREAD'S OWN, not the `lambda: None`
        // the prelude passed. Passing a no-op meant the spawned thread did
        // nothing and the work happened in a `join`-time drain queue, which is
        // why a body never observed anything another thread did.
        function(
            "__py_thread_start",
            vec![param("thread", None)],
            vec![if_stmt(
                unary_not(field_of(ident("thread"), "_started")),
                vec![
                    assign(index(ident("thread"), str_lit("_started")), bool_lit(true)),
                    // ⛔ A NULLARY CLOSURE, not the bare target. The worker
                    // chunk `emit_thread_start_with` builds invokes what it is
                    // handed with ZERO arguments
                    // (`emit_direct_invoke_chunk(worker, 0)`), so a target
                    // declared `def worker(n)` ran with `n` undefined and
                    // `Thread(target=w, args=(i,))` produced `[nan, nan, nan,
                    // nan]`. The closure carries `_args` across the spawn.
                    expr_stmt(call_global(
                        "__py_thread_start_common",
                        vec![
                            ident("thread"),
                            call_global("__py_thread_runner", vec![ident("thread")]),
                        ],
                    )),
                ],
            )],
        ),
        // `join` is the real one too. `_done` is stamped after it returns so
        // `is_alive()` answers correctly once the thread has been waited on.
        function(
            "__py_thread_join",
            vec![param("thread", None), param("timeout", Some(null()))],
            vec![
                expr_stmt(call_global(
                    "__py_thread_join_common",
                    vec![ident("thread")],
                )),
                assign(index(ident("thread"), str_lit("_done")), bool_lit(true)),
                ret(null()),
            ],
        ),
        global_assign(
            "__py_main_thread",
            new("Thread", vec![null(), null(), str_lit("MainThread")]),
        ),
        function("current_thread", vec![], vec![ret(ident("__py_main_thread"))]),
        function("main_thread", vec![], vec![ret(ident("__py_main_thread"))]),
        function(
            "enumerate",
            vec![],
            vec![ret(list_of(vec![ident("__py_main_thread")]))],
        ),
        stub_fn("active_count", num(1.0)),
        stub_fn("get_ident", num(1.0)),
        stub_fn("stack_size", num(8388608.0)),
        function("excepthook", any_args(), vec![ret(null())]),
        // ⛔ THE WALKER REWRITES `threading.Thread(...)` TO THIS.
        // `walker.rs:14851` names `__py_thread_make` directly so it can pass a
        // `target_name` the source never wrote (the pending queue orders
        // "producer" threads first). Deleting the prelude without porting this
        // left every `Thread(...)` construction calling `undefined`.
        thread_factory("__py_thread_make", "Thread"),
    ]
}

/// `__py_<x>_make(...)` — the factory the walker rewrites a construction to.
/// `Thread` and `Process` differ only in which class is built.
pub(crate) fn thread_factory(fn_name: &str, class_name: &'static str) -> Statement {
    function(
        fn_name,
        vec![
            param("group", Some(null())),
            param("target", Some(null())),
            param("name", Some(null())),
            param("args", Some(list_of(vec![]))),
            param("kwargs", Some(null())),
            param("daemon", Some(null())),
            param("target_name", Some(str_lit(""))),
        ],
        vec![
            assign(
                ident("__t"),
                new(
                    class_name,
                    vec![
                        ident("group"),
                        ident("target"),
                        ident("name"),
                        ident("args"),
                        ident("kwargs"),
                        ident("daemon"),
                    ],
                ),
            ),
            assign(
                index(ident("__t"), str_lit("_target_name")),
                ident("target_name"),
            ),
            ret(ident("__t")),
        ],
    )
}

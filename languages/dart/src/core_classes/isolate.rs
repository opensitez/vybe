//! `dart:isolate`, as classes over the shared channel model.
//!
//! A `ReceivePort` IS a channel: `sendPort.send` is `ChanOp::Send`, `first`
//! is the BLOCKING `ChanOp::Recv` (futex wait-slice), so a message sent from
//! a spawned isolate — `Isolate.spawn` rides go's exact lowering,
//! `common:threading.task_run` → `wasi:threads.thread-spawn` — wakes the
//! awaiting receiver through machinery every CSP language already exercises.
//! Nothing here is isolate-private: the channel value crosses threads because
//! heap objects cross threads, and the blocking receive is the same helper
//! chunk go's `<-ch` links.
//!
//! Two dart-isms sit BESIDE the channel rather than in it, in a state record
//! `{closed, handler}` shared by the port and its `SendPort`:
//! - dart drops sends to a CLOSED port where go panics, so `send` tests
//!   `closed` before it ever touches the channel (and `close` never calls
//!   `ChanOp::Close` — the go-spec panic-on-double-close is not dart's
//!   contract);
//! - `listen` dispatches through a HANDLER: registering one first drains
//!   whatever the channel buffered (a spawned isolate may send before the
//!   listener exists), then `send` invokes it directly instead of buffering.
//!
//! The colliding member names — `first`, `take`, `listen`,
//! `asBroadcastStream` are receiver-blind `[value_methods]` rows — are
//! declared here under `__vybe*` spellings; `walker.rs` renames the member on
//! receivers it can type as ports. Declaring the real names as METHODS would
//! put them into the flat `defined_class_methods` set that every untyped
//! receiver consults, diverting `stream.listen` / `iterable.take` everywhere
//! (the StringBuffer `length` lesson: measured 0/50 across dart slices).

use super::builders::*;
use vybe_ast::{
    Argument, ChanOp, ExprKind, Expression, ObjectProperty, Statement, StmtKind,
};

const CHAN: &str = "_vybeChan";
const STATE: &str = "_vybeState";
const SEND_PORT: &str = "_vybeSendPort";
const BYTES: &str = "_vybeBytes";

fn null_lit() -> Expression {
    Expression::null()
}

/// `{closed: false, handler: null}` — the port/send-port shared state.
fn state_record() -> Expression {
    Expression::new(ExprKind::Object(vec![
        ObjectProperty::KeyValue {
            key: Expression::string("closed"),
            value: bool_lit(false),
        },
        ObjectProperty::KeyValue {
            key: Expression::string("handler"),
            value: null_lit(),
        },
    ]))
}

/// An effectively unbounded channel: dart ports never refuse a send.
fn chan_new() -> Expression {
    Expression::new(ExprKind::Chan(ChanOp::New {
        capacity: Some(Box::new(int_lit(1_000_000_000))),
        zero: Box::new(null_lit()),
    }))
}

fn chan_send(channel: Expression, value: Expression) -> Expression {
    Expression::new(ExprKind::Chan(ChanOp::Send {
        channel: Box::new(channel),
        value: Box::new(value),
    }))
}

fn chan_recv(channel: Expression) -> Expression {
    Expression::new(ExprKind::Chan(ChanOp::Recv(Box::new(channel))))
}

fn chan_try_recv(channel: Expression) -> Expression {
    Expression::new(ExprKind::Chan(ChanOp::TryRecv(Box::new(channel))))
}

fn index(object: Expression, i: i64) -> Expression {
    Expression::new(ExprKind::Index {
        object: Box::new(object),
        index: Box::new(int_lit(i)),
        null_safe: false,
    })
}

fn index_expr(object: Expression, i: Expression) -> Expression {
    Expression::new(ExprKind::Index {
        object: Box::new(object),
        index: Box::new(i),
        null_safe: false,
    })
}

fn while_stmt(cond: Expression, body: Vec<Statement>) -> Statement {
    Statement::with_span(
        StmtKind::While {
            cond,
            body,
            else_body: None,
        },
        span(),
    )
}

fn new_of(class: &str, args: Vec<Expression>) -> Expression {
    Expression::new(ExprKind::New {
        class: Box::new(ident(class)),
        args: args.into_iter().map(Argument::positional).collect(),
    })
}

fn call_fn(callee: Expression, args: Vec<Expression>) -> Expression {
    Expression::new(ExprKind::Call {
        callee: Box::new(callee),
        args: args.into_iter().map(Argument::positional).collect(),
        optional: false,
    })
}

fn ret_none() -> Statement {
    Statement::with_span(StmtKind::Return(None), span())
}

fn this_expr() -> Expression {
    Expression::new(ExprKind::This)
}

/// `class Capability {}` — an unforgeable token. Identity `==` and the
/// identity hash are exactly dart's contract for one.
pub(super) fn capability() -> Statement {
    class("Capability", vec![constructor(vec![], vec![])])
}

/// `class SendPort { … }` — the send half. Constructed only by `ReceivePort`
/// (and `Isolate.controlPort`); both ends share `chan` and `state`.
pub(super) fn send_port() -> Statement {
    let members = vec![
        field(CHAN, "dynamic", null_lit()),
        field(STATE, "dynamic", null_lit()),
        constructor(
            vec![param("chan", None, None), param("state", None, None)],
            vec![
                set_this(CHAN, ident("chan")),
                set_this(STATE, ident("state")),
            ],
        ),
        method(
            "send",
            vec![param("message", None, None)],
            Some("void"),
            vec![
                // Dropped, not panicked: dart's closed-port contract.
                if_stmt(
                    binary(
                        vybe_ast::BinOp::Eq,
                        field_of(this_field(STATE), "closed"),
                        bool_lit(true),
                    ),
                    vec![ret_none()],
                ),
                if_stmt(
                    binary(
                        vybe_ast::BinOp::NotEq,
                        field_of(this_field(STATE), "handler"),
                        null_lit(),
                    ),
                    vec![
                        expr_stmt(call_fn(
                            field_of(this_field(STATE), "handler"),
                            vec![ident("message")],
                        )),
                        ret_none(),
                    ],
                ),
                expr_stmt(chan_send(this_field(CHAN), ident("message"))),
            ],
        ),
    ];
    class("SendPort", members)
}

/// The members `ReceivePort` and `RawReceivePort` share. `with_handler` adds
/// the raw port's construction-time handler parameter.
fn port_members(with_handler: bool) -> Vec<vybe_ast::ClassMember> {
    let mut ctor_body = vec![
        set_this(CHAN, chan_new()),
        set_this(STATE, state_record()),
        set_this(
            SEND_PORT,
            new_of("SendPort", vec![this_field(CHAN), this_field(STATE)]),
        ),
    ];
    let ctor_params = if with_handler {
        ctor_body.push(if_stmt(
            binary(vybe_ast::BinOp::NotEq, ident("handler"), null_lit()),
            vec![assign(
                field_of(this_field(STATE), "handler"),
                ident("handler"),
            )],
        ));
        vec![param("handler", None, Some(null_lit()))]
    } else {
        vec![]
    };

    vec![
        field(CHAN, "dynamic", null_lit()),
        field(STATE, "dynamic", null_lit()),
        field(SEND_PORT, "dynamic", null_lit()),
        constructor(ctor_params, ctor_body),
        // A zero-arg METHOD, force-called by name (`is_dart_zero_arg_getter`):
        // a property getter's body cannot see `this` — the Uri classes
        // measured this first (see `core_classes/uri.rs`).
        method(
            "sendPort",
            vec![],
            Some("SendPort"),
            vec![ret(this_field(SEND_PORT))],
        ),
        method(
            "close",
            vec![],
            Some("void"),
            vec![
                assign(field_of(this_field(STATE), "closed"), bool_lit(true)),
                assign(field_of(this_field(STATE), "handler"), null_lit()),
            ],
        ),
        // `port.first` — the walker renames the member on port-typed
        // receivers, so the `[value_methods] first` row keeps owning lists.
        // A zero-arg METHOD, not a getter: `first` is on the walker's
        // zero-arg-getter list, so the read arrives as a CALL (see
        // `ZERO_ARG_GETTER_NAMES` in mod.rs), and the rename pass emits a
        // call for the plain-read spelling too.
        method(
            "__vybeFirst",
            vec![],
            Some("dynamic"),
            vec![ret(chan_recv(this_field(CHAN)))],
        ),
        method(
            "__vybeListen",
            vec![param("handler", None, None)],
            Some("void"),
            vec![
                // Drain what buffered before the listener existed …
                local("__pair", chan_try_recv(this_field(CHAN))),
                while_stmt(
                    binary(
                        vybe_ast::BinOp::Eq,
                        index(ident("__pair"), 1),
                        bool_lit(true),
                    ),
                    vec![
                        expr_stmt(call_fn(ident("handler"), vec![index(ident("__pair"), 0)])),
                        assign(ident("__pair"), chan_try_recv(this_field(CHAN))),
                    ],
                ),
                // … then let `send` dispatch directly.
                assign(field_of(this_field(STATE), "handler"), ident("handler")),
            ],
        ),
        method(
            "__vybeTake",
            vec![param("count", None, None)],
            Some("List"),
            vec![
                local("__out", empty_list()),
                local("__i", int_lit(0)),
                while_stmt(
                    binary(vybe_ast::BinOp::Lt, ident("__i"), ident("count")),
                    vec![
                        expr_stmt(call_member(
                            ident("__out"),
                            "add",
                            vec![chan_recv(this_field(CHAN))],
                        )),
                        assign(
                            ident("__i"),
                            binary(vybe_ast::BinOp::Add, ident("__i"), int_lit(1)),
                        ),
                    ],
                ),
                ret(ident("__out")),
            ],
        ),
        method(
            "__vybeAsBroadcast",
            vec![],
            Some("dynamic"),
            vec![ret(this_expr())],
        ),
    ]
}

/// The ports declare `Stream` as an INTERFACE (dart's `ReceivePort` extends
/// `Stream`): `classes.rs` emits an instanceof chain for each declared
/// interface name, so `port is Stream` answers true without a `Stream` class
/// existing anywhere.
fn port_class(name: &str, members: Vec<vybe_ast::ClassMember>) -> Statement {
    Statement::with_span(
        StmtKind::ClassDecl {
            name: name.to_string(),
            parents: Vec::new(),
            interfaces: vec!["Stream".to_string()],
            members,
            modifiers: vybe_ast::ClassModifiers::default(),
            decorators: vec![],
        },
        span(),
    )
}

pub(super) fn receive_port() -> Statement {
    port_class("ReceivePort", port_members(false))
}

pub(super) fn raw_receive_port() -> Statement {
    port_class("RawReceivePort", port_members(true))
}

/// `class Isolate { … }`. `Isolate.spawn` is rewritten by the walker into
/// `Isolate(__dart_isolate_spawn(() { entry(message); }))` — the builtin is
/// go's `__go_spawn` row (`common:threading.task_run`), so the entry runs on
/// a real spawned thread and its sends rendezvous through the channel.
///
/// `pause`/`resume`/`kill` are capability theatre by design: a wasi thread
/// cannot be suspended or destroyed from outside, and the corpus asserts the
/// CONTROL-SURFACE shapes (a `Capability` comes back, the message still
/// arrives), not preemption.
pub(super) fn isolate() -> Statement {
    let members = vec![
        field("_vybeTask", "dynamic", null_lit()),
        constructor(
            vec![param("task", None, Some(null_lit()))],
            vec![set_this("_vybeTask", ident("task"))],
        ),
        method(
            "kill",
            vec![param("priority", None, Some(null_lit()))],
            Some("void"),
            vec![],
        ),
        method(
            "pause",
            vec![param("resumeCapability", None, Some(null_lit()))],
            Some("Capability"),
            vec![ret(new_of("Capability", vec![]))],
        ),
        method(
            "resume",
            vec![param("resumeCapability", None, None)],
            Some("void"),
            vec![],
        ),
        method(
            "ping",
            vec![
                param("responsePort", None, None),
                param("response", None, Some(null_lit())),
            ],
            Some("void"),
            vec![expr_stmt(call_member(
                ident("responsePort"),
                "send",
                vec![ident("response")],
            ))],
        ),
        // The spawned entry has, in this model, already run to completion by
        // the time a listener can be added — sending the response right away
        // IS the at-exit delivery order the corpus observes.
        method(
            "addOnExitListener",
            vec![
                param("responsePort", None, None),
                param("response", None, Some(null_lit())),
            ],
            Some("void"),
            vec![expr_stmt(call_member(
                ident("responsePort"),
                "send",
                vec![ident("response")],
            ))],
        ),
        method(
            "removeOnExitListener",
            vec![param("responsePort", None, None)],
            Some("void"),
            vec![],
        ),
        method(
            "addErrorListener",
            vec![param("responsePort", None, None)],
            Some("void"),
            vec![],
        ),
        method(
            "removeErrorListener",
            vec![param("responsePort", None, None)],
            Some("void"),
            vec![],
        ),
        // Zero-arg methods, force-called by name — see the port note above.
        // The main isolate's debug name IS "main" (dart VM default).
        method(
            "debugName",
            vec![],
            Some("String"),
            vec![ret(str_lit("main"))],
        ),
        method(
            "controlPort",
            vec![],
            Some("SendPort"),
            vec![ret(new_of("SendPort", vec![chan_new(), state_record()]))],
        ),
        method(
            "terminateCapability",
            vec![],
            Some("Capability"),
            vec![ret(new_of("Capability", vec![]))],
        ),
        method(
            "pauseCapability",
            vec![],
            Some("Capability"),
            vec![ret(new_of("Capability", vec![]))],
        ),
    ];
    class("Isolate", members)
}

/// `dart:async`'s `Completer`, over the same channel model: `complete` is a
/// send, the `future` getter is the blocking receive, so `await c.future`
/// after a `complete` (from any thread) reads the completed value. Lives in
/// this file because the channel vocabulary is here, not because it is an
/// isolate type.
pub(super) fn completer() -> Statement {
    let members = vec![
        field(CHAN, "dynamic", null_lit()),
        constructor(vec![], vec![set_this(CHAN, chan_new())]),
        method(
            "complete",
            vec![param("value", None, Some(null_lit()))],
            Some("void"),
            vec![expr_stmt(chan_send(this_field(CHAN), ident("value")))],
        ),
        // Zero-arg method + force-call, like the port getters.
        method(
            "future",
            vec![],
            Some("dynamic"),
            vec![ret(chan_recv(this_field(CHAN)))],
        ),
    ];
    class("Completer", members)
}

/// `class TransferableTypedData { … }` — the walker rewrites
/// `TransferableTypedData.fromList(lists)` to a construction. The byte
/// concatenation is spelled with plain loops and `[]`/`.add`, so it lowers
/// through the shared collection machinery; `materialize` answers a real
/// `DataView` over the bytes, which is what `ByteData` aliases to, so the
/// test-side `getUint8`/`getFloat32`/`lengthInBytes` reads dispatch on it
/// unchanged.
pub(super) fn transferable_typed_data() -> Statement {
    let members = vec![
        field(BYTES, "dynamic", null_lit()),
        field("_vybeMaterialized", "bool", bool_lit(false)),
        constructor(
            vec![param("lists", None, None)],
            vec![
                // One `addAll` per list, not a per-byte loop: a 100KB
                // Uint8List concatenated element-wise in the interpreter
                // timed out the runner (measured).
                local("__out", empty_list()),
                for_in(
                    "__l",
                    ident("lists"),
                    vec![expr_stmt(call_member(
                        ident("__out"),
                        "addAll",
                        vec![ident("__l")],
                    ))],
                ),
                set_this(BYTES, ident("__out")),
            ],
        ),
        // `materialize()` answers a ByteBuffer (the ArrayBuffer under a dart
        // name — `.asByteData()` and `lengthInBytes` are walker renames), and
        // it MOVES: a second call throws ArgumentError (dart 3.10.4,
        // measured).
        method(
            "materialize",
            vec![],
            Some("dynamic"),
            vec![
                if_stmt(
                    binary(
                        vybe_ast::BinOp::Eq,
                        this_field("_vybeMaterialized"),
                        bool_lit(true),
                    ),
                    vec![Statement::with_span(
                        StmtKind::Throw {
                            expr: Some(new_of(
                                "ArgumentError",
                                vec![str_lit("Cannot materialize twice")],
                            )),
                            cause: None,
                        },
                        span(),
                    )],
                ),
                set_this("_vybeMaterialized", bool_lit(true)),
                ret(field_of(
                    new_of("Uint8Array", vec![this_field(BYTES)]),
                    "buffer",
                )),
            ],
        ),
    ];
    class("TransferableTypedData", members)
}

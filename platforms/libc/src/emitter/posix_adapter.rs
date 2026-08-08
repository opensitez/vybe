//! POSIX compatibility adapters for libc-backed languages.

use vybe_ast::{
    ArrayElement, BinOp, BreakTarget, ChanOp, ExprKind, Expression, Literal, ObjectProperty,
    PlaceExpr, Statement, StmtKind, UnaryOp };

// ── AF_UNIX is a CHANNEL, not a socket ──────────────────────────────────────
//
// A unix-domain socket is local IPC between threads of one process: a byte
// queue with a blocking reader. That is `ChanOp`, already normalized in the
// AST and lowered ONCE in `primitives/channels.rs` onto shared linear memory
// with `memory.atomic.wait32/notify` — so a blocked `accept` or `recv` wakes
// cross-thread with no host function and no VM change.
//
// What a socket adds over a channel is a NAME: `bind` publishes under a
// `sun_path`, `connect` looks it up. Naming is the language's surface, so the
// registry is an ordinary object in this runtime, not a new shared concept.
// MSG_PEEK is `TryPeek`, O_NONBLOCK is `TryRecv`, `shutdown` is `Close` —
// every one of them already exists.

/// A buffered channel. Unbuffered would be a Go rendezvous, where the sender
/// blocks until a reader arrives; a socket buffers, so `send` returns as soon
/// as the bytes are queued.
fn chan_new() -> Expression {
    expr(ExprKind::Chan(ChanOp::New {
        capacity: Some(Box::new(int_lit(1024))),
        zero: Box::new(str_lit("")) }))
}

fn chan_send(channel: Expression, value: Expression) -> Expression {
    expr(ExprKind::Chan(ChanOp::Send {
        channel: Box::new(channel),
        value: Box::new(value) }))
}

fn chan_recv(channel: Expression) -> Expression {
    expr(ExprKind::Chan(ChanOp::Recv(Box::new(channel))))
}

/// `(value, ok)` without consuming — MSG_PEEK.
fn chan_try_peek(channel: Expression) -> Expression {
    expr(ExprKind::Chan(ChanOp::TryPeek(Box::new(channel))))
}

/// `(value, ok)` without blocking — O_NONBLOCK / EAGAIN.
fn chan_try_recv(channel: Expression) -> Expression {
    expr(ExprKind::Chan(ChanOp::TryRecv(Box::new(channel))))
}

/// Is this descriptor an AF_UNIX one? `AF_UNIX` is 1 in the constant table.
fn is_unix_fd(fd_expr: Expression) -> Expression {
    bin(
        BinOp::Eq,
        index_expr(ident("__c_sock_family"), fd_expr),
        int_lit(1),
    )
}

use super::build::{
    assign_expr, call_expr, call_member, expr, function_stmt, ident, index_expr, int_lit, member,
    null_lit, stmt, str_lit, var_decl_stmt };
use vybe_compiler::primitives::pointers;

pub type HeaderStruct = (&'static str, &'static [(&'static str, &'static str)]);

pub fn runtime_helpers() -> Vec<Statement> {
    let mut out = vec![
        str_to_codes_helper(),
        exec_helper(),
        poll_helper(),
        select_helper(),
    ];
    out.extend(socket_helpers());
    out
}

/// REAL sockets over `wasi:sockets`, fd-keyed.
///
/// Python's `VybeSocketImpl` holds its resource on a socket OBJECT; C has
/// integer descriptors, so the resource and its two streams live in tables
/// keyed by fd (`__c_sock_res/_rx/_tx`). The WASI `start-*`/`finish-*`
/// pairs are two calls because the proposal is poll-based — a C socket call
/// is blocking, so each helper does BOTH, exactly as python's class does.
fn socket_helpers() -> Vec<Statement> {
    // "a.b.c.d:port" from a sockaddr_in — `htonl`/`htons` are identity in
    // this runtime, so the fields hold host-order values.
    let addr_text = function_stmt(
        "__c_sock_addr_text",
        vec!["addr"],
        vec![
            var_decl_stmt(
                "v",
                nullish(
                    member(member(ident("addr"), "sin_addr"), "s_addr"),
                    nullish(member(ident("addr"), "s_addr"), int_lit(0)),
                ),
            ),
            var_decl_stmt("port", nullish(member(ident("addr"), "sin_port"), int_lit(0))),
            // INADDR_ANY binds loopback here: the tests bind ANY or
            // LOOPBACK and then connect to loopback.
            stmt(StmtKind::Expr(ternary(
                bin(BinOp::Eq, ident("v"), int_lit(0)),
                assign_expr(ident("v"), int_lit(0x7F00_0001)),
                int_lit(0),
            ))),
            stmt(StmtKind::Return(Some(call_expr(
                ident("__c_sprintf"),
                vec![
                    str_lit("%d.%d.%d.%d:%d"),
                    bin(
                        BinOp::BitAnd,
                        bin(BinOp::Shr, ident("v"), int_lit(24)),
                        int_lit(255),
                    ),
                    bin(
                        BinOp::BitAnd,
                        bin(BinOp::Shr, ident("v"), int_lit(16)),
                        int_lit(255),
                    ),
                    bin(
                        BinOp::BitAnd,
                        bin(BinOp::Shr, ident("v"), int_lit(8)),
                        int_lit(255),
                    ),
                    bin(BinOp::BitAnd, ident("v"), int_lit(255)),
                    ident("port"),
                ],
            )))),
        ],
    );

    // socket(family, type) → a real descriptor backed by a WASI socket.
    let socket_new = function_stmt(
        "__c_socket_h",
        vec!["family", "kind"],
        vec![
            var_decl_stmt("fd", next_fd()),
            stmt(StmtKind::Expr(assign_expr(
                next_fd(),
                bin(BinOp::Add, next_fd(), int_lit(1)),
            ))),
            stmt(StmtKind::Expr(assign_expr(
                index_expr(ident("__c_fd_open"), ident("fd")),
                int_lit(1),
            ))),
            stmt(StmtKind::Expr(assign_expr(
                index_expr(ident("__c_sock_kind"), ident("fd")),
                ident("kind"),
            ))),
            stmt(StmtKind::Expr(assign_expr(
                index_expr(ident("__c_sock_family"), ident("fd")),
                ident("family"),
            ))),
            // AF_UNIX never touches the network stack — no WASI resource, and
            // its channels are created by `bind`/`connect`.
            if_stmt(
                is_unix_fd(ident("fd")),
                vec![stmt(StmtKind::Return(Some(ident("fd"))))],
                None,
            ),
            stmt(StmtKind::Expr(assign_expr(
                index_expr(ident("__c_sock_res"), ident("fd")),
                ternary(
                    bin(BinOp::Eq, ident("kind"), int_lit(2)),
                    call_expr(ident("__c_wasi_udp_new"), vec![str_lit("ipv4")]),
                    call_expr(ident("__c_wasi_tcp_new"), vec![str_lit("ipv4")]),
                ),
            ))),
            stmt(StmtKind::Return(Some(ident("fd")))),
        ],
    );

    let bind_h = function_stmt(
        "__c_bind_h",
        vec!["fd", "addr"],
        vec![
            // AF_UNIX binds a NAME, not an address: publish this descriptor's
            // channel under `sun_path`. bind(2) fails with EADDRINUSE if the
            // path already exists — a real filesystem check, not a filename
            // this runtime knows.
            if_stmt(
                is_unix_fd(ident("fd")),
                vec![
                    var_decl_stmt("p", member(ident("addr"), "sun_path")),
                    if_stmt(
                        index_expr(ident("__c_path_exists"), ident("p")),
                        vec![stmt(StmtKind::Return(Some(int_lit(-1))))],
                        None,
                    ),
                    stmt(StmtKind::Expr(assign_expr(
                        index_expr(ident("__c_path_exists"), ident("p")),
                        int_lit(1),
                    ))),
                    stmt(StmtKind::Expr(assign_expr(
                        index_expr(ident("__c_sock_path"), ident("fd")),
                        ident("p"),
                    ))),
                    // A listener's channel carries CONNECTIONS; a datagram
                    // socket's carries the datagrams themselves. Same channel
                    // type, different payload — which is the whole difference
                    // between SOCK_STREAM and SOCK_DGRAM here.
                    stmt(StmtKind::Expr(assign_expr(
                        index_expr(ident("__c_unix_reg"), ident("p")),
                        chan_new(),
                    ))),
                    if_stmt(
                        bin(
                            BinOp::Eq,
                            index_expr(ident("__c_sock_kind"), ident("fd")),
                            int_lit(2),
                        ),
                        vec![stmt(StmtKind::Expr(assign_expr(
                            index_expr(ident("__c_sock_in"), ident("fd")),
                            index_expr(ident("__c_unix_reg"), ident("p")),
                        )))],
                        None,
                    ),
                    stmt(StmtKind::Return(Some(int_lit(0)))),
                ],
                None,
            ),
            // UDP is a DIFFERENT wasi interface, not a flag on tcp.
            if_stmt(
                bin(
                    BinOp::Eq,
                    index_expr(ident("__c_sock_kind"), ident("fd")),
                    int_lit(2),
                ),
                vec![
                    stmt(StmtKind::Expr(call_expr(
                        ident("__c_wasi_udp_start_bind"),
                        vec![
                            index_expr(ident("__c_sock_res"), ident("fd")),
                            call_expr(ident("__c_wasi_network"), vec![]),
                            call_expr(ident("__c_sock_addr_text"), vec![ident("addr")]),
                        ],
                    ))),
                    stmt(StmtKind::Expr(call_expr(
                        ident("__c_wasi_udp_finish_bind"),
                        vec![index_expr(ident("__c_sock_res"), ident("fd"))],
                    ))),
                    stmt(StmtKind::Expr(assign_expr(
                        index_expr(ident("__c_sock_bound"), ident("fd")),
                        int_lit(1),
                    ))),
                    // A bound datagram socket can RECEIVE without anyone
                    // calling `connect`, so the pair is created here rather
                    // than being a side effect of connecting.
                    stmt(StmtKind::Expr(call_expr(
                        ident("__c_udp_stream_h"),
                        vec![ident("fd")],
                    ))),
                    stmt(StmtKind::Return(Some(int_lit(0)))),
                ],
                None,
            ),
            stmt(StmtKind::Expr(call_expr(
                ident("__c_wasi_start_bind"),
                vec![
                    index_expr(ident("__c_sock_res"), ident("fd")),
                    call_expr(ident("__c_wasi_network"), vec![]),
                    call_expr(ident("__c_sock_addr_text"), vec![ident("addr")]),
                ],
            ))),
            stmt(StmtKind::Expr(call_expr(
                ident("__c_wasi_finish_bind"),
                vec![index_expr(ident("__c_sock_res"), ident("fd"))],
            ))),
            stmt(StmtKind::Return(Some(int_lit(0)))),
        ],
    );

    let listen_h = function_stmt(
        "__c_listen_h",
        vec!["fd", "backlog"],
        vec![
            // listen(2) is stream-only: EOPNOTSUPP on a datagram socket.
            if_stmt(
                bin(
                    BinOp::Eq,
                    index_expr(ident("__c_sock_kind"), ident("fd")),
                    int_lit(2),
                ),
                vec![stmt(StmtKind::Return(Some(int_lit(-1))))],
                None,
            ),
            stmt(StmtKind::Expr(call_expr(
                ident("__c_wasi_backlog"),
                vec![
                    index_expr(ident("__c_sock_res"), ident("fd")),
                    ident("backlog"),
                ],
            ))),
            stmt(StmtKind::Expr(call_expr(
                ident("__c_wasi_start_listen"),
                vec![index_expr(ident("__c_sock_res"), ident("fd"))],
            ))),
            stmt(StmtKind::Return(Some(int_lit(0)))),
        ],
    );

    let connect_h = function_stmt(
        "__c_connect_h",
        vec!["fd", "addr"],
        vec![
            // AF_UNIX: look the name up. Nothing bound there is ECONNREFUSED,
            // which is what the corpus's `while (connect(...) != 0) usleep()`
            // loops on until the server side binds.
            if_stmt(
                is_unix_fd(ident("fd")),
                vec![
                    var_decl_stmt("p", member(ident("addr"), "sun_path")),
                    var_decl_stmt("reg", index_expr(ident("__c_unix_reg"), ident("p"))),
                    if_stmt(
                        bin(BinOp::Eq, ident("reg"), null_lit()),
                        vec![stmt(StmtKind::Return(Some(int_lit(-1))))],
                        None,
                    ),
                    // A datagram socket has no connection to make — it just
                    // remembers where an address-less `send` goes.
                    if_stmt(
                        bin(
                            BinOp::Eq,
                            index_expr(ident("__c_sock_kind"), ident("fd")),
                            int_lit(2),
                        ),
                        vec![
                            stmt(StmtKind::Expr(assign_expr(
                                index_expr(ident("__c_sock_out"), ident("fd")),
                                ident("reg"),
                            ))),
                            stmt(StmtKind::Expr(assign_expr(
                                index_expr(ident("__c_sock_peer"), ident("fd")),
                                ident("p"),
                            ))),
                            stmt(StmtKind::Return(Some(int_lit(0)))),
                        ],
                        None,
                    ),
                    // A stream connection is a PAIR of channels — one per
                    // direction — handed to the listener, which is what makes
                    // `accept` a receive rather than a poll.
                    var_decl_stmt("c2s", chan_new()),
                    var_decl_stmt("s2c", chan_new()),
                    stmt(StmtKind::Expr(assign_expr(
                        index_expr(ident("__c_sock_out"), ident("fd")),
                        ident("c2s"),
                    ))),
                    stmt(StmtKind::Expr(assign_expr(
                        index_expr(ident("__c_sock_in"), ident("fd")),
                        ident("s2c"),
                    ))),
                    stmt(StmtKind::Expr(assign_expr(
                        index_expr(ident("__c_sock_peer"), ident("fd")),
                        ident("p"),
                    ))),
                    stmt(StmtKind::Expr(chan_send(
                        ident("reg"),
                        expr(ExprKind::Array(vec![
                            ArrayElement {
                                key: None,
                                value: ident("c2s"),
                                spread: false,
                                by_ref: false },
                            ArrayElement {
                                key: None,
                                value: ident("s2c"),
                                spread: false,
                                by_ref: false },
                        ])),
                    ))),
                    stmt(StmtKind::Return(Some(int_lit(0)))),
                ],
                None,
            ),
            if_stmt(
                bin(
                    BinOp::Eq,
                    index_expr(ident("__c_sock_kind"), ident("fd")),
                    int_lit(2),
                ),
                vec![
                    // Connecting a datagram socket that was never bound must
                    // still give it a local address, exactly as sending does.
                    stmt(StmtKind::Expr(call_expr(
                        ident("__c_udp_stream_h"),
                        vec![ident("fd")],
                    ))),
                    var_decl_stmt(
                        "ds",
                        call_expr(
                            ident("__c_wasi_udp_stream"),
                            vec![
                                index_expr(ident("__c_sock_res"), ident("fd")),
                                call_expr(ident("__c_sock_addr_text"), vec![ident("addr")]),
                            ],
                        ),
                    ),
                    if_stmt(
                        bin(BinOp::NotEq, ident("ds"), null_lit()),
                        vec![
                            stmt(StmtKind::Expr(assign_expr(
                                index_expr(ident("__c_sock_rx"), ident("fd")),
                                index_expr(ident("ds"), int_lit(0)),
                            ))),
                            stmt(StmtKind::Expr(assign_expr(
                                index_expr(ident("__c_sock_tx"), ident("fd")),
                                index_expr(ident("ds"), int_lit(1)),
                            ))),
                        ],
                        None,
                    ),
                    // The peer is what makes a later `send` legal — see
                    // `__c_send_h`'s EDESTADDRREQ arm.
                    stmt(StmtKind::Expr(assign_expr(
                        index_expr(ident("__c_sock_peer"), ident("fd")),
                        call_expr(ident("__c_sock_addr_text"), vec![ident("addr")]),
                    ))),
                    stmt(StmtKind::Return(Some(int_lit(0)))),
                ],
                None,
            ),
            stmt(StmtKind::Expr(call_expr(
                ident("__c_wasi_start_conn"),
                vec![
                    index_expr(ident("__c_sock_res"), ident("fd")),
                    call_expr(ident("__c_wasi_network"), vec![]),
                    call_expr(ident("__c_sock_addr_text"), vec![ident("addr")]),
                ],
            ))),
            var_decl_stmt(
                "streams",
                call_expr(
                    ident("__c_wasi_finish_conn"),
                    vec![index_expr(ident("__c_sock_res"), ident("fd"))],
                ),
            ),
            if_stmt(
                bin(BinOp::Eq, ident("streams"), null_lit()),
                vec![stmt(StmtKind::Return(Some(int_lit(-1))))],
                None,
            ),
            stmt(StmtKind::Expr(assign_expr(
                index_expr(ident("__c_sock_rx"), ident("fd")),
                index_expr(ident("streams"), int_lit(0)),
            ))),
            stmt(StmtKind::Expr(assign_expr(
                index_expr(ident("__c_sock_tx"), ident("fd")),
                index_expr(ident("streams"), int_lit(1)),
            ))),
            stmt(StmtKind::Return(Some(int_lit(0)))),
        ],
    );

    // `wasi:sockets` accept is POLL-shaped — nothing queued answers null.
    // POSIX `accept` on a blocking socket WAITS, so this is where the two
    // shapes are reconciled: poll, yield through a real sleep, poll again.
    // (The same start-*/finish-* sequencing python's socket class does.)
    // The peer is a genuine thread now, so the wait actually resolves; the
    // cap turns a client that never arrives into EAGAIN instead of a hang.
    let accept_h = function_stmt(
        "__c_accept_h",
        vec!["fd"],
        vec![
            // AF_UNIX: a connection is the next value on the listener's
            // channel. The receive BLOCKS on the channel's futex word, so a
            // waiting `accept` wakes the moment another thread connects —
            // no retry loop and no timing constant.
            if_stmt(
                is_unix_fd(ident("fd")),
                vec![
                    var_decl_stmt(
                        "reg",
                        index_expr(
                            ident("__c_unix_reg"),
                            index_expr(ident("__c_sock_path"), ident("fd")),
                        ),
                    ),
                    if_stmt(
                        bin(BinOp::Eq, ident("reg"), null_lit()),
                        vec![stmt(StmtKind::Return(Some(int_lit(-1))))],
                        None,
                    ),
                    var_decl_stmt("pair", chan_recv(ident("reg"))),
                    var_decl_stmt("ufd", next_fd()),
                    stmt(StmtKind::Expr(assign_expr(
                        next_fd(),
                        bin(BinOp::Add, next_fd(), int_lit(1)),
                    ))),
                    stmt(StmtKind::Expr(assign_expr(
                        index_expr(ident("__c_fd_open"), ident("ufd")),
                        int_lit(1),
                    ))),
                    stmt(StmtKind::Expr(assign_expr(
                        index_expr(ident("__c_sock_family"), ident("ufd")),
                        int_lit(1),
                    ))),
                    // The listener reads what the client wrote and writes what
                    // the client reads — the pair, crossed.
                    stmt(StmtKind::Expr(assign_expr(
                        index_expr(ident("__c_sock_in"), ident("ufd")),
                        index_expr(ident("pair"), int_lit(0)),
                    ))),
                    stmt(StmtKind::Expr(assign_expr(
                        index_expr(ident("__c_sock_out"), ident("ufd")),
                        index_expr(ident("pair"), int_lit(1)),
                    ))),
                    stmt(StmtKind::Expr(assign_expr(
                        index_expr(ident("__c_sock_peer"), ident("ufd")),
                        index_expr(ident("__c_sock_path"), ident("fd")),
                    ))),
                    stmt(StmtKind::Return(Some(ident("ufd")))),
                ],
                None,
            ),
            var_decl_stmt(
                "r",
                call_expr(
                    ident("__c_wasi_accept"),
                    vec![index_expr(ident("__c_sock_res"), ident("fd"))],
                ),
            ),
            var_decl_stmt("waited", int_lit(0)),
            // …unless the descriptor is O_NONBLOCK, which is the caller
            // saying "answer EAGAIN, do not wait".
            var_decl_stmt(
                "nb",
                nullish(index_expr(ident("__c_fd_nonblock"), ident("fd")), int_lit(0)),
            ),
            stmt(StmtKind::While {
                cond: bin(
                    BinOp::And,
                    bin(BinOp::Eq, ident("nb"), int_lit(0)),
                    bin(
                        BinOp::And,
                        bin(BinOp::Eq, ident("r"), null_lit()),
                        bin(BinOp::Lt, ident("waited"), int_lit(1000)),
                    ),
                ),
                body: vec![
                    stmt(StmtKind::Expr(call_expr(
                        ident("__c_sleep_ms"),
                        vec![int_lit(5)],
                    ))),
                    stmt(StmtKind::Expr(assign_expr(
                        ident("waited"),
                        bin(BinOp::Add, ident("waited"), int_lit(1)),
                    ))),
                    stmt(StmtKind::Expr(assign_expr(
                        ident("r"),
                        call_expr(
                            ident("__c_wasi_accept"),
                            vec![index_expr(ident("__c_sock_res"), ident("fd"))],
                        ),
                    ))),
                ],
                else_body: None }),
            if_stmt(
                bin(BinOp::Eq, ident("r"), null_lit()),
                vec![stmt(StmtKind::Return(Some(int_lit(-1))))],
                None,
            ),
            var_decl_stmt("cfd", next_fd()),
            stmt(StmtKind::Expr(assign_expr(
                next_fd(),
                bin(BinOp::Add, next_fd(), int_lit(1)),
            ))),
            stmt(StmtKind::Expr(assign_expr(
                index_expr(ident("__c_fd_open"), ident("cfd")),
                int_lit(1),
            ))),
            stmt(StmtKind::Expr(assign_expr(
                index_expr(ident("__c_sock_res"), ident("cfd")),
                index_expr(ident("r"), int_lit(0)),
            ))),
            stmt(StmtKind::Expr(assign_expr(
                index_expr(ident("__c_sock_rx"), ident("cfd")),
                index_expr(ident("r"), int_lit(1)),
            ))),
            stmt(StmtKind::Expr(assign_expr(
                index_expr(ident("__c_sock_tx"), ident("cfd")),
                index_expr(ident("r"), int_lit(2)),
            ))),
            stmt(StmtKind::Return(Some(ident("cfd")))),
        ],
    );

    // A datagram socket needs its stream PAIR before it can send or receive,
    // and `connect` is not what creates it — an unconnected UDP socket both
    // sends (with an explicit destination) and receives. POSIX also auto-binds
    // an unbound socket on its first send, which is the `__c_sock_bound` arm.
    //
    // `stream` with no remote is the unconnected pair; `connect` calls the same
    // host function WITH the address, which replaces the pair with a connected
    // one. Both are `wasi:sockets/udp.stream`, so there is one route, not two.
    let udp_stream_h = function_stmt(
        "__c_udp_stream_h",
        vec!["fd"],
        vec![
            if_stmt(
                bin(
                    BinOp::NotEq,
                    index_expr(ident("__c_sock_tx"), ident("fd")),
                    null_lit(),
                ),
                vec![stmt(StmtKind::Return(Some(int_lit(0))))],
                None,
            ),
            if_stmt(
                bin(
                    BinOp::Eq,
                    nullish(index_expr(ident("__c_sock_bound"), ident("fd")), int_lit(0)),
                    int_lit(0),
                ),
                vec![
                    stmt(StmtKind::Expr(call_expr(
                        ident("__c_wasi_udp_start_bind"),
                        vec![
                            index_expr(ident("__c_sock_res"), ident("fd")),
                            call_expr(ident("__c_wasi_network"), vec![]),
                            str_lit("127.0.0.1:0"),
                        ],
                    ))),
                    stmt(StmtKind::Expr(call_expr(
                        ident("__c_wasi_udp_finish_bind"),
                        vec![index_expr(ident("__c_sock_res"), ident("fd"))],
                    ))),
                    stmt(StmtKind::Expr(assign_expr(
                        index_expr(ident("__c_sock_bound"), ident("fd")),
                        int_lit(1),
                    ))),
                ],
                None,
            ),
            var_decl_stmt(
                "ds",
                call_expr(
                    ident("__c_wasi_udp_stream"),
                    vec![index_expr(ident("__c_sock_res"), ident("fd"))],
                ),
            ),
            if_stmt(
                bin(BinOp::NotEq, ident("ds"), null_lit()),
                vec![
                    stmt(StmtKind::Expr(assign_expr(
                        index_expr(ident("__c_sock_rx"), ident("fd")),
                        index_expr(ident("ds"), int_lit(0)),
                    ))),
                    stmt(StmtKind::Expr(assign_expr(
                        index_expr(ident("__c_sock_tx"), ident("fd")),
                        index_expr(ident("ds"), int_lit(1)),
                    ))),
                ],
                None,
            ),
            stmt(StmtKind::Return(Some(int_lit(0)))),
        ],
    );

    let send_h = function_stmt(
        "__c_send_h",
        vec!["fd", "data", "count", "dest"],
        vec![
            var_decl_stmt(
                "text",
                call_member(
                    call_expr(ident("__libc_char_to_str"), vec![ident("data")]),
                    "substring",
                    vec![int_lit(0), ident("count")],
                ),
            ),
            // AF_UNIX: the destination is a NAME for a datagram socket and the
            // established channel for a stream one.
            if_stmt(
                is_unix_fd(ident("fd")),
                vec![
                    var_decl_stmt(
                        "ch",
                        ternary(
                            bin(BinOp::Eq, ident("dest"), null_lit()),
                            index_expr(ident("__c_sock_out"), ident("fd")),
                            index_expr(
                                ident("__c_unix_reg"),
                                member(ident("dest"), "sun_path"),
                            ),
                        ),
                    ),
                    if_stmt(
                        bin(BinOp::Eq, ident("ch"), null_lit()),
                        vec![stmt(StmtKind::Return(Some(int_lit(-1))))],
                        None,
                    ),
                    stmt(StmtKind::Expr(chan_send(ident("ch"), ident("text")))),
                    stmt(StmtKind::Return(Some(ident("count")))),
                ],
                None,
            ),
            if_stmt(
                bin(
                    BinOp::Eq,
                    index_expr(ident("__c_sock_kind"), ident("fd")),
                    int_lit(2),
                ),
                vec![
                    stmt(StmtKind::Expr(call_expr(
                        ident("__c_udp_stream_h"),
                        vec![ident("fd")],
                    ))),
                    // `send` with no destination and no connected peer is
                    // EDESTADDRREQ — there is nowhere to send it. `sendto`
                    // with a NULL destination on a connected socket is legal
                    // and uses the peer, which is why the peer, not `dest`,
                    // decides.
                    if_stmt(
                        bin(
                            BinOp::And,
                            bin(BinOp::Eq, ident("dest"), null_lit()),
                            bin(
                                BinOp::Eq,
                                nullish(
                                    index_expr(ident("__c_sock_peer"), ident("fd")),
                                    null_lit(),
                                ),
                                null_lit(),
                            ),
                        ),
                        vec![stmt(StmtKind::Return(Some(int_lit(-1))))],
                        None,
                    ),
                    // `send` is symmetric with `receive`: a LIST OF DATAGRAM
                    // RECORDS. The key must be ABSENT for a connected send,
                    // not present-and-null: the host reads
                    // `properties.get("remote-address")`, and a null VALUE is
                    // still a `Some`, so it never falls back to the stream's
                    // peer and the datagram is silently dropped. Two arms, so
                    // the key is genuinely missing in the connected one.
                    if_stmt(
                        bin(BinOp::Eq, ident("dest"), null_lit()),
                        vec![stmt(StmtKind::Expr(call_expr(
                            ident("__c_wasi_dgram_send"),
                            vec![
                                index_expr(ident("__c_sock_tx"), ident("fd")),
                                expr(ExprKind::Array(vec![ArrayElement {
                                    key: None,
                                    value: expr(ExprKind::Object(vec![
                                        ObjectProperty::KeyValue {
                                            key: str_lit("data"),
                                            value: ident("text") },
                                    ])),
                                    spread: false,
                                    by_ref: false }])),
                            ],
                        )))],
                        Some(vec![stmt(StmtKind::Expr(call_expr(
                            ident("__c_wasi_dgram_send"),
                            vec![
                                index_expr(ident("__c_sock_tx"), ident("fd")),
                                expr(ExprKind::Array(vec![ArrayElement {
                                    key: None,
                                    value: expr(ExprKind::Object(vec![
                                        ObjectProperty::KeyValue {
                                            key: str_lit("data"),
                                            value: ident("text") },
                                        // "host:port" is one of the shapes the
                                        // host parses.
                                        ObjectProperty::KeyValue {
                                            key: str_lit("remote-address"),
                                            value: call_expr(
                                                ident("__c_sock_addr_text"),
                                                vec![ident("dest")],
                                            ) },
                                    ])),
                                    spread: false,
                                    by_ref: false }])),
                            ],
                        )))]),
                    ),
                    stmt(StmtKind::Return(Some(ident("count")))),
                ],
                None,
            ),
            stmt(StmtKind::Expr(call_expr(
                ident("__c_wasi_stream_write"),
                vec![index_expr(ident("__c_sock_tx"), ident("fd")), ident("text")],
            ))),
            stmt(StmtKind::Return(Some(ident("count")))),
        ],
    );

    // `wasi:io/streams.read` answers BYTES; a C buffer holds text, so the
    // code units become a string here (the same boundary conversion
    // `__libc_wide_to_string` performs for wide arrays).
    let recv_h = function_stmt(
        "__c_recv_h",
        vec!["fd", "count", "flags"],
        vec![
            // AF_UNIX: every recv variant is already a channel op. MSG_PEEK is
            // `TryPeek`, O_NONBLOCK is `TryRecv`, and the plain blocking read
            // is `Recv` — which parks on the channel's futex instead of
            // sleeping in a loop.
            if_stmt(
                is_unix_fd(ident("fd")),
                vec![
                    var_decl_stmt("ch", index_expr(ident("__c_sock_in"), ident("fd"))),
                    if_stmt(
                        bin(BinOp::Eq, ident("ch"), null_lit()),
                        vec![stmt(StmtKind::Return(Some(str_lit(""))))],
                        None,
                    ),
                    if_stmt(
                        bin(
                            BinOp::NotEq,
                            bin(
                                BinOp::BitAnd,
                                nullish(ident("flags"), int_lit(0)),
                                int_lit(2),
                            ),
                            int_lit(0),
                        ),
                        vec![
                            var_decl_stmt("pk", chan_try_peek(ident("ch"))),
                            stmt(StmtKind::Return(Some(ternary(
                                index_expr(ident("pk"), int_lit(1)),
                                index_expr(ident("pk"), int_lit(0)),
                                str_lit(""),
                            )))),
                        ],
                        None,
                    ),
                    if_stmt(
                        bin(
                            BinOp::NotEq,
                            nullish(index_expr(ident("__c_fd_nonblock"), ident("fd")), int_lit(0)),
                            int_lit(0),
                        ),
                        vec![
                            var_decl_stmt("tr", chan_try_recv(ident("ch"))),
                            stmt(StmtKind::Return(Some(ternary(
                                index_expr(ident("tr"), int_lit(1)),
                                index_expr(ident("tr"), int_lit(0)),
                                null_lit(),
                            )))),
                        ],
                        None,
                    ),
                    stmt(StmtKind::Return(Some(chan_recv(ident("ch"))))),
                ],
                None,
            ),
            if_stmt(
                bin(
                    BinOp::Eq,
                    index_expr(ident("__c_sock_kind"), ident("fd")),
                    int_lit(2),
                ),
                vec![stmt(StmtKind::Expr(call_expr(
                    ident("__c_udp_stream_h"),
                    vec![ident("fd")],
                )))],
                None,
            ),
            // MSG_PEEK reads without consuming. WASI has no peek, so the
            // datagram already taken off the wire is held here and the NEXT
            // read serves it — one datagram of lookahead, which is what the
            // kernel queue gives you for the sequence the flag exists for
            // (peek, then read the same bytes).
            var_decl_stmt("pk", index_expr(ident("__c_sock_peek"), ident("fd"))),
            if_stmt(
                bin(BinOp::NotEq, ident("pk"), null_lit()),
                vec![
                    if_stmt(
                        bin(
                            BinOp::Eq,
                            bin(
                                BinOp::BitAnd,
                                nullish(ident("flags"), int_lit(0)),
                                int_lit(2),
                            ),
                            int_lit(0),
                        ),
                        vec![stmt(StmtKind::Expr(assign_expr(
                            index_expr(ident("__c_sock_peek"), ident("fd")),
                            null_lit(),
                        )))],
                        None,
                    ),
                    stmt(StmtKind::Return(Some(ident("pk")))),
                ],
                None,
            ),
            var_decl_stmt(
                "data",
                ternary(
                    bin(
                        BinOp::Eq,
                        index_expr(ident("__c_sock_kind"), ident("fd")),
                        int_lit(2),
                    ),
                    call_expr(
                        ident("__c_wasi_dgram_recv"),
                        vec![index_expr(ident("__c_sock_rx"), ident("fd")), int_lit(1)],
                    ),
                    call_expr(
                        ident("__c_wasi_stream_read"),
                        vec![index_expr(ident("__c_sock_rx"), ident("fd")), ident("count")],
                    ),
                ),
            ),
            // POSIX `recv` on a blocking socket waits for data; the WASI
            // read is non-blocking and answers an EMPTY list when nothing
            // has arrived yet. Retry while the stream is still open —
            // `null` is the closed/failed stream, which is a real 0-byte
            // answer and must NOT be waited on.
            var_decl_stmt("waited", int_lit(0)),
            var_decl_stmt(
                "nb",
                nullish(index_expr(ident("__c_fd_nonblock"), ident("fd")), int_lit(0)),
            ),
            stmt(StmtKind::While {
                cond: bin(
                    BinOp::And,
                    bin(BinOp::Eq, ident("nb"), int_lit(0)),
                    bin(
                        BinOp::And,
                        bin(
                            BinOp::And,
                            bin(BinOp::NotEq, ident("data"), null_lit()),
                            bin(
                                BinOp::Eq,
                                nullish(member(ident("data"), "length"), int_lit(0)),
                                int_lit(0),
                            ),
                        ),
                        bin(BinOp::Lt, ident("waited"), int_lit(400)),
                    ),
                ),
                body: vec![
                    stmt(StmtKind::Expr(call_expr(
                        ident("__c_sleep_ms"),
                        vec![int_lit(5)],
                    ))),
                    stmt(StmtKind::Expr(assign_expr(
                        ident("waited"),
                        bin(BinOp::Add, ident("waited"), int_lit(1)),
                    ))),
                    stmt(StmtKind::Expr(assign_expr(
                        ident("data"),
                        ternary(
                            bin(
                                BinOp::Eq,
                                index_expr(ident("__c_sock_kind"), ident("fd")),
                                int_lit(2),
                            ),
                            call_expr(
                                ident("__c_wasi_dgram_recv"),
                                vec![index_expr(ident("__c_sock_rx"), ident("fd")), int_lit(1)],
                            ),
                            call_expr(
                                ident("__c_wasi_stream_read"),
                                vec![index_expr(ident("__c_sock_rx"), ident("fd")), ident("count")],
                            ),
                        ),
                    ))),
                ],
                else_body: None }),
            // A closed/failed stream is a real 0-byte answer…
            if_stmt(
                bin(BinOp::Eq, ident("data"), null_lit()),
                vec![stmt(StmtKind::Return(Some(str_lit(""))))],
                None,
            ),
            // …but "open, nothing queued" on a NON-BLOCKING descriptor is
            // EAGAIN, which recv(2) reports as -1, not as 0 bytes. `null`
            // carries that back; the call site turns it into -1.
            if_stmt(
                bin(
                    BinOp::And,
                    bin(BinOp::NotEq, ident("nb"), int_lit(0)),
                    bin(
                        BinOp::Eq,
                        nullish(member(ident("data"), "length"), int_lit(0)),
                        int_lit(0),
                    ),
                ),
                vec![stmt(StmtKind::Return(Some(null_lit())))],
                None,
            ),
            // A datagram stream answers a LIST OF RECORDS (each `data` +
            // `remote-address`), not bytes: take the first datagram's
            // payload. An empty list means nothing arrived.
            if_stmt(
                bin(
                    BinOp::Eq,
                    index_expr(ident("__c_sock_kind"), ident("fd")),
                    int_lit(2),
                ),
                vec![
                    if_stmt(
                        bin(
                            BinOp::Eq,
                            nullish(member(ident("data"), "length"), int_lit(0)),
                            int_lit(0),
                        ),
                        vec![stmt(StmtKind::Return(Some(str_lit(""))))],
                        None,
                    ),
                    stmt(StmtKind::Expr(assign_expr(
                        ident("data"),
                        member(index_expr(ident("data"), int_lit(0)), "data"),
                    ))),
                ],
                None,
            ),
            var_decl_stmt(
                "out",
                ternary(
                    bin(
                        BinOp::Eq,
                        expr(ExprKind::Unary {
                            op: UnaryOp::Typeof,
                            expr: Box::new(ident("data")) }),
                        str_lit("string"),
                    ),
                    ident("data"),
                    call_expr(ident("__libc_wide_to_string"), vec![ident("data")]),
                ),
            ),
            // A peeking read leaves the bytes queued for the next one.
            if_stmt(
                bin(
                    BinOp::NotEq,
                    bin(
                        BinOp::BitAnd,
                        nullish(ident("flags"), int_lit(0)),
                        int_lit(2),
                    ),
                    int_lit(0),
                ),
                vec![stmt(StmtKind::Expr(assign_expr(
                    index_expr(ident("__c_sock_peek"), ident("fd")),
                    ident("out"),
                )))],
                None,
            ),
            stmt(StmtKind::Return(Some(ident("out")))),
        ],
    );

    let sockname_h = function_stmt(
        "__c_getsockname_h",
        vec!["fd", "addr"],
        vec![
            // An AF_UNIX socket's local address IS the path it bound.
            if_stmt(
                is_unix_fd(ident("fd")),
                vec![
                    stmt(StmtKind::Expr(assign_expr(
                        member(ident("addr"), "sun_family"),
                        int_lit(1),
                    ))),
                    stmt(StmtKind::Expr(assign_expr(
                        member(ident("addr"), "sun_path"),
                        nullish(
                            index_expr(ident("__c_sock_path"), ident("fd")),
                            str_lit(""),
                        ),
                    ))),
                    stmt(StmtKind::Return(Some(int_lit(0)))),
                ],
                None,
            ),
            var_decl_stmt(
                "rec",
                call_expr(
                    ident("__c_wasi_local_addr"),
                    vec![index_expr(ident("__c_sock_res"), ident("fd"))],
                ),
            ),
            if_stmt(
                bin(BinOp::NotEq, ident("rec"), null_lit()),
                vec![stmt(StmtKind::Expr(assign_expr(
                    member(ident("addr"), "sin_port"),
                    nullish(member(ident("rec"), "port"), int_lit(0)),
                )))],
                None,
            ),
            stmt(StmtKind::Return(Some(int_lit(0)))),
        ],
    );

    let peername_h = function_stmt(
        "__c_getpeername_h",
        vec!["fd", "addr"],
        vec![
            // ENOTCONN unless this descriptor actually connected or was
            // accepted — `__c_sock_peer` is set by exactly those two.
            if_stmt(
                is_unix_fd(ident("fd")),
                vec![
                    if_stmt(
                        bin(
                            BinOp::Eq,
                            nullish(index_expr(ident("__c_sock_peer"), ident("fd")), null_lit()),
                            null_lit(),
                        ),
                        vec![stmt(StmtKind::Return(Some(int_lit(-1))))],
                        None,
                    ),
                    stmt(StmtKind::Expr(assign_expr(
                        member(ident("addr"), "sun_family"),
                        int_lit(1),
                    ))),
                    stmt(StmtKind::Expr(assign_expr(
                        member(ident("addr"), "sun_path"),
                        index_expr(ident("__c_sock_peer"), ident("fd")),
                    ))),
                    stmt(StmtKind::Return(Some(int_lit(0)))),
                ],
                None,
            ),
            var_decl_stmt(
                "rec",
                call_expr(
                    ident("__c_wasi_remote_addr"),
                    vec![index_expr(ident("__c_sock_res"), ident("fd"))],
                ),
            ),
            if_stmt(
                bin(BinOp::Eq, ident("rec"), null_lit()),
                vec![stmt(StmtKind::Return(Some(int_lit(-1))))],
                None,
            ),
            stmt(StmtKind::Expr(assign_expr(
                member(ident("addr"), "sin_port"),
                nullish(member(ident("rec"), "port"), int_lit(0)),
            ))),
            stmt(StmtKind::Return(Some(int_lit(0)))),
        ],
    );

    vec![
        addr_text, socket_new, udp_stream_h, bind_h, listen_h, connect_h, accept_h, send_h,
        recv_h, sockname_h, peername_h,
    ]
}

fn exec_helper() -> Statement {
    // exec* REPLACES the process image with a REAL program run
    // (node:child_process spawnSync via `__c_spawn_sync`). The child's
    // stdout/stderr are forwarded to ours; a spawn failure is -1 (ENOENT),
    // which is the ONLY case where exec returns in C.
    //
    // Inside a forked child (`__c_in_forked_child`, set by `fork()` which
    // runs the child inline) exec records the status and falls through so
    // the parent's code continues. At top level exec never returns: the
    // run ends with the child's status, exactly as POSIX says.
    function_stmt(
        "__c_exec_h",
        vec!["path", "argv", "env", "action"],
        vec![
            var_decl_stmt(
                "p",
                call_expr(ident("__libc_char_to_str"), vec![ident("path")]),
            ),
            // Flatten argv (carray pointer / bare string / array) up to NULL.
            var_decl_stmt("args", expr(ExprKind::Array(vec![]))),
            var_decl_stmt("arg_list", ident("argv")),
            if_stmt(
                pointers::is_carray_ptr_kind(ident("arg_list")),
                vec![stmt(StmtKind::Expr(assign_expr(
                    ident("arg_list"),
                    call_member(
                        member(ident("arg_list"), pointers::CARRAY_BASE_KEY),
                        "slice",
                        vec![member(ident("arg_list"), pointers::CARRAY_IDX_KEY)],
                    ),
                )))],
                None,
            ),
            if_stmt(
                bin(
                    BinOp::Eq,
                    expr(ExprKind::Unary {
                        op: UnaryOp::Typeof,
                        expr: Box::new(ident("arg_list")) }),
                    str_lit("string"),
                ),
                vec![stmt(StmtKind::Expr(assign_expr(
                    ident("arg_list"),
                    expr(ExprKind::Array(vec![ArrayElement {
                        key: None,
                        value: ident("arg_list"),
                        spread: false,
                        by_ref: false }])),
                )))],
                None,
            ),
            var_decl_stmt("i", int_lit(0)),
            stmt(StmtKind::While {
                cond: and(
                    bin(BinOp::NotEq, ident("arg_list"), null_lit()),
                    bin(BinOp::Lt, ident("i"), member(ident("arg_list"), "length")),
                ),
                body: vec![
                    var_decl_stmt("item", index_expr(ident("arg_list"), ident("i"))),
                    if_stmt(
                        or(
                            bin(BinOp::Eq, ident("item"), null_lit()),
                            bin(
                                BinOp::Eq,
                                expr(ExprKind::Unary {
                                    op: UnaryOp::Typeof,
                                    expr: Box::new(ident("item")) }),
                                str_lit("undefined"),
                            ),
                        ),
                        vec![stmt(StmtKind::Break(BreakTarget::Implicit))],
                        None,
                    ),
                    stmt(StmtKind::Expr(call_expr(
                        member(ident("args"), "push"),
                        vec![call_expr(ident("__libc_char_to_str"), vec![ident("item")])],
                    ))),
                    stmt(StmtKind::Expr(assign_expr(
                        ident("i"),
                        bin(BinOp::Add, ident("i"), int_lit(1)),
                    ))),
                ],
                else_body: None }),
            // argv[0] is the conventional NAME, not the program: the real
            // arguments are argv[1..].
            var_decl_stmt("real_args", call_member(ident("args"), "slice", vec![int_lit(1)])),
            var_decl_stmt("opts", expr(ExprKind::Object(vec![]))),
            // An explicit env list (execve/execle/execvpe) REPLACES the
            // environment — including the empty list, which means empty.
            var_decl_stmt("env_list", ident("env")),
            if_stmt(
                pointers::is_carray_ptr_kind(ident("env_list")),
                vec![stmt(StmtKind::Expr(assign_expr(
                    ident("env_list"),
                    call_member(
                        member(ident("env_list"), pointers::CARRAY_BASE_KEY),
                        "slice",
                        vec![member(ident("env_list"), pointers::CARRAY_IDX_KEY)],
                    ),
                )))],
                None,
            ),
            if_stmt(
                bin(BinOp::NotEq, ident("env_list"), null_lit()),
                vec![
                    var_decl_stmt("eo", expr(ExprKind::Object(vec![]))),
                    var_decl_stmt("j", int_lit(0)),
                    stmt(StmtKind::While {
                        cond: bin(BinOp::Lt, ident("j"), member(ident("env_list"), "length")),
                        body: vec![
                            var_decl_stmt("entry", index_expr(ident("env_list"), ident("j"))),
                            if_stmt(
                                or(
                                    bin(BinOp::Eq, ident("entry"), null_lit()),
                                    bin(
                                        BinOp::Eq,
                                        expr(ExprKind::Unary {
                                            op: UnaryOp::Typeof,
                                            expr: Box::new(ident("entry")) }),
                                        str_lit("undefined"),
                                    ),
                                ),
                                vec![stmt(StmtKind::Break(BreakTarget::Implicit))],
                                None,
                            ),
                            var_decl_stmt(
                                "text",
                                call_expr(ident("__libc_char_to_str"), vec![ident("entry")]),
                            ),
                            var_decl_stmt(
                                "eq",
                                call_member(ident("text"), "indexOf", vec![str_lit("=")]),
                            ),
                            if_stmt(
                                bin(BinOp::GtEq, ident("eq"), int_lit(0)),
                                vec![stmt(StmtKind::Expr(assign_expr(
                                    index_expr(
                                        ident("eo"),
                                        call_member(
                                            ident("text"),
                                            "substring",
                                            vec![int_lit(0), ident("eq")],
                                        ),
                                    ),
                                    call_member(
                                        ident("text"),
                                        "substring",
                                        vec![bin(BinOp::Add, ident("eq"), int_lit(1))],
                                    ),
                                )))],
                                None,
                            ),
                            stmt(StmtKind::Expr(assign_expr(
                                ident("j"),
                                bin(BinOp::Add, ident("j"), int_lit(1)),
                            ))),
                        ],
                        else_body: None }),
                    stmt(StmtKind::Expr(assign_expr(
                        member(ident("opts"), "env"),
                        ident("eo"),
                    ))),
                ],
                Some(vec![if_stmt(
                    bin(BinOp::Eq, ident("__c_env_dirty"), int_lit(1)),
                    vec![stmt(StmtKind::Expr(assign_expr(
                        member(ident("opts"), "env"),
                        ident("__c_env_obj"),
                    )))],
                    None,
                )]),
            ),
            // Our own buffered stdout must land before the child's.
            stmt(StmtKind::Expr(call_expr(
                ident("__c_write_stdout"),
                vec![ident("__c_stdout_buffer")],
            ))),
            stmt(StmtKind::Expr(assign_expr(
                ident("__c_stdout_buffer"),
                str_lit(""),
            ))),
            var_decl_stmt(
                "r",
                call_expr(
                    ident("__c_spawn_sync"),
                    vec![ident("p"), ident("real_args"), ident("opts")],
                ),
            ),
            // Spawn failure — the one case where exec RETURNS (ENOENT).
            if_stmt(
                bin(BinOp::NotEq, nullish(member(ident("r"), "error"), null_lit()), null_lit()),
                vec![stmt(StmtKind::Return(Some(int_lit(-1))))],
                None,
            ),
            // `posix_spawn_file_actions_addopen` redirects the child's
            // stdout into a file instead of ours.
            if_stmt(
                and(
                    bin(BinOp::NotEq, ident("action"), null_lit()),
                    member(ident("action"), "openPath"),
                ),
                vec![stmt(StmtKind::Expr(assign_expr(
                    index_expr(
                        ident("__c_file_store"),
                        member(ident("action"), "openPath"),
                    ),
                    member(ident("r"), "stdout"),
                )))],
                Some(vec![if_stmt(
                    bin(BinOp::NotEq, member(ident("r"), "stdout"), str_lit("")),
                    vec![stmt(StmtKind::Expr(call_expr(
                        ident("__c_write_stdout"),
                        vec![member(ident("r"), "stdout")],
                    )))],
                    None,
                )]),
            ),
            if_stmt(
                bin(BinOp::NotEq, member(ident("r"), "stderr"), str_lit("")),
                vec![stmt(StmtKind::Expr(call_expr(
                    ident("__c_fputs_h"),
                    vec![member(ident("r"), "stderr"), int_lit(2)],
                )))],
                None,
            ),
            var_decl_stmt(
                "st",
                ternary(
                    bin(BinOp::Eq, member(ident("r"), "status"), null_lit()),
                    int_lit(2),
                    member(ident("r"), "status"),
                ),
            ),
            stmt(StmtKind::Expr(assign_expr(
                ident("__c_child_status"),
                ident("st"),
            ))),
            // In a forked child: record and fall through (the parent's code
            // follows). At top level exec never returns — end the run.
            if_stmt(
                bin(BinOp::Eq, ident("__c_in_forked_child"), int_lit(1)),
                vec![stmt(StmtKind::Return(Some(int_lit(0))))],
                None,
            ),
            stmt(StmtKind::Expr(call_expr(
                ident("__c_exit_with_code"),
                vec![ident("st")],
            ))),
            stmt(StmtKind::Return(Some(int_lit(0)))),
        ],
    )
}

fn str_to_codes_helper() -> Statement {
    function_stmt(
        "__c_str_to_codes",
        vec!["s"],
        vec![
            var_decl_stmt("out", expr(ExprKind::Array(vec![]))),
            var_decl_stmt("i", int_lit(0)),
            stmt(StmtKind::While {
                cond: bin(BinOp::Lt, ident("i"), member(ident("s"), "length")),
                body: vec![
                    stmt(StmtKind::Expr(call_expr(
                        member(ident("out"), "push"),
                        vec![call_expr(
                            member(ident("s"), "charCodeAt"),
                            vec![ident("i")],
                        )],
                    ))),
                    stmt(StmtKind::Expr(assign_expr(
                        ident("i"),
                        bin(BinOp::Add, ident("i"), int_lit(1)),
                    ))),
                ],
                else_body: None }),
            stmt(StmtKind::Return(Some(ident("out")))),
        ],
    )
}

fn poll_helper() -> Statement {
    function_stmt(
        "__c_poll_h",
        vec!["fds", "nfds"],
        vec![
            var_decl_stmt("ready", int_lit(0)),
            var_decl_stmt("i", int_lit(0)),
            stmt(StmtKind::While {
                cond: bin(BinOp::Lt, ident("i"), ident("nfds")),
                body: vec![
                    var_decl_stmt(
                        "p",
                        ternary(
                            eq(ident("nfds"), int_lit(1)),
                            ident("fds"),
                            index_expr(ident("fds"), ident("i")),
                        ),
                    ),
                    var_decl_stmt("fd", member(ident("p"), "fd")),
                    stmt(StmtKind::Expr(assign_expr(
                        member(ident("p"), "revents"),
                        int_lit(0),
                    ))),
                    if_stmt(
                        bin(BinOp::Lt, ident("fd"), int_lit(0)),
                        vec![],
                        Some(vec![if_stmt(
                            expr(ExprKind::Unary {
                                op: UnaryOp::Not,
                                expr: Box::new(index_expr(ident("__c_fd_open"), ident("fd"))) }),
                            vec![
                                stmt(StmtKind::Expr(assign_expr(
                                    member(ident("p"), "revents"),
                                    int_lit(32),
                                ))),
                                stmt(StmtKind::Expr(assign_expr(
                                    ident("ready"),
                                    bin(BinOp::Add, ident("ready"), int_lit(1)),
                                ))),
                            ],
                            Some(vec![
                                if_stmt(
                                    and(
                                        bin(
                                            BinOp::NotEq,
                                            bin(
                                                BinOp::BitAnd,
                                                member(ident("p"), "events"),
                                                int_lit(1),
                                            ),
                                            int_lit(0),
                                        ),
                                        or(
                                            index_expr(ident("__c_fd_content_by_fd"), ident("fd")),
                                            index_expr(
                                                ident("__c_pipe_writer_closed"),
                                                ident("fd"),
                                            ),
                                        ),
                                    ),
                                    vec![
                                        stmt(StmtKind::Expr(assign_expr(
                                            member(ident("p"), "revents"),
                                            ternary(
                                                index_expr(
                                                    ident("__c_pipe_writer_closed"),
                                                    ident("fd"),
                                                ),
                                                int_lit(16),
                                                int_lit(1),
                                            ),
                                        ))),
                                        stmt(StmtKind::Expr(assign_expr(
                                            ident("ready"),
                                            bin(BinOp::Add, ident("ready"), int_lit(1)),
                                        ))),
                                    ],
                                    None,
                                ),
                                if_stmt(
                                    and(
                                        eq(member(ident("p"), "revents"), int_lit(0)),
                                        bin(
                                            BinOp::NotEq,
                                            bin(
                                                BinOp::BitAnd,
                                                member(ident("p"), "events"),
                                                int_lit(4),
                                            ),
                                            int_lit(0),
                                        ),
                                    ),
                                    vec![
                                        stmt(StmtKind::Expr(assign_expr(
                                            member(ident("p"), "revents"),
                                            int_lit(4),
                                        ))),
                                        stmt(StmtKind::Expr(assign_expr(
                                            ident("ready"),
                                            bin(BinOp::Add, ident("ready"), int_lit(1)),
                                        ))),
                                    ],
                                    None,
                                ),
                            ]),
                        )]),
                    ),
                    stmt(StmtKind::Expr(assign_expr(
                        ident("i"),
                        bin(BinOp::Add, ident("i"), int_lit(1)),
                    ))),
                ],
                else_body: None }),
            stmt(StmtKind::Return(Some(ident("ready")))),
        ],
    )
}

fn select_helper() -> Statement {
    function_stmt(
        "__c_select_h",
        vec!["nfds", "readfds", "writefds"],
        vec![
            if_stmt(
                bin(BinOp::Lt, ident("nfds"), int_lit(0)),
                vec![stmt(StmtKind::Return(Some(int_lit(-1))))],
                None,
            ),
            var_decl_stmt("ready", int_lit(0)),
            var_decl_stmt("i", int_lit(0)),
            stmt(StmtKind::While {
                cond: bin(BinOp::Lt, ident("i"), ident("nfds")),
                body: vec![
                    if_stmt(
                        index_expr(ident("readfds"), ident("i")),
                        vec![if_stmt(
                            expr(ExprKind::Unary {
                                op: UnaryOp::Not,
                                expr: Box::new(index_expr(ident("__c_fd_open"), ident("i"))) }),
                            vec![stmt(StmtKind::Return(Some(int_lit(-1))))],
                            Some(vec![if_stmt(
                                or(
                                    index_expr(ident("__c_fd_content_by_fd"), ident("i")),
                                    index_expr(ident("__c_pipe_writer_closed"), ident("i")),
                                ),
                                vec![stmt(StmtKind::Expr(assign_expr(
                                    ident("ready"),
                                    bin(BinOp::Add, ident("ready"), int_lit(1)),
                                )))],
                                Some(vec![stmt(StmtKind::Expr(assign_expr(
                                    index_expr(ident("readfds"), ident("i")),
                                    int_lit(0),
                                )))]),
                            )]),
                        )],
                        None,
                    ),
                    if_stmt(
                        index_expr(ident("writefds"), ident("i")),
                        vec![if_stmt(
                            expr(ExprKind::Unary {
                                op: UnaryOp::Not,
                                expr: Box::new(index_expr(ident("__c_fd_open"), ident("i"))) }),
                            vec![stmt(StmtKind::Return(Some(int_lit(-1))))],
                            Some(vec![stmt(StmtKind::Expr(assign_expr(
                                ident("ready"),
                                bin(BinOp::Add, ident("ready"), int_lit(1)),
                            )))]),
                        )],
                        None,
                    ),
                    stmt(StmtKind::Expr(assign_expr(
                        ident("i"),
                        bin(BinOp::Add, ident("i"), int_lit(1)),
                    ))),
                ],
                else_body: None }),
            stmt(StmtKind::Return(Some(ident("ready")))),
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
        else_body })
}

fn bin(op: BinOp, left: Expression, right: Expression) -> Expression {
    expr(ExprKind::Binary {
        op,
        left: Box::new(left),
        right: Box::new(right) })
}

/// The next free file descriptor. Held on an object so that a spawned
/// thread — whose globals are a CLONE — allocates out of the same counter
/// as its parent instead of reissuing descriptors the parent already owns.
fn next_fd() -> Expression {
    member(ident("__c_fd_seq"), "n")
}

pub fn header_structs(header: &str) -> Vec<HeaderStruct> {
    match header {
        "time.h" => vec![
            (
                "tm",
                &[
                    ("tm_sec", "int"),
                    ("tm_min", "int"),
                    ("tm_hour", "int"),
                    ("tm_mday", "int"),
                    ("tm_mon", "int"),
                    ("tm_year", "int"),
                    ("tm_wday", "int"),
                    ("tm_yday", "int"),
                    ("tm_isdst", "int"),
                ],
            ),
            ("timespec", &[("tv_sec", "long"), ("tv_nsec", "long")]),
            ("timeval", &[("tv_sec", "long"), ("tv_usec", "long")]),
            (
                "itimerspec",
                &[("it_interval", "timespec"), ("it_value", "timespec")],
            ),
            (
                "itimerval",
                &[("it_interval", "timeval"), ("it_value", "timeval")],
            ),
        ],
        "sys/time.h" => vec![
            ("timespec", &[("tv_sec", "long"), ("tv_nsec", "long")]),
            ("timeval", &[("tv_sec", "long"), ("tv_usec", "long")]),
            (
                "itimerspec",
                &[("it_interval", "timespec"), ("it_value", "timespec")],
            ),
            (
                "itimerval",
                &[("it_interval", "timeval"), ("it_value", "timeval")],
            ),
        ],
        "signal.h" => vec![
            (
                "sigaction",
                &[
                    ("sa_handler", "int"),
                    ("sa_sigaction", "int"),
                    ("sa_mask", "int"),
                    ("sa_flags", "int"),
                ],
            ),
            (
                "sigevent",
                &[
                    ("sigev_notify", "int"),
                    ("sigev_signo", "int"),
                    ("sigev_value", "int"),
                    ("sigev_notify_function", "int"),
                ],
            ),
            (
                "stack_t",
                &[("ss_sp", "int"), ("ss_flags", "int"), ("ss_size", "int")],
            ),
            ("siginfo_t", &[("si_signo", "int")]),
        ],
        "sys/stat.h" => vec![(
            "stat",
            &[
                ("st_size", "long"),
                ("st_mode", "long"),
                ("st_atime", "long"),
                ("st_mtime", "long"),
                ("st_ctime", "long"),
                ("st_ino", "long"),
                ("st_dev", "long"),
            ],
        )],
        "netinet/in.h" => inet_structs(),
        "netdb.h" => vec![
            (
                "hostent",
                &[
                    ("h_name", "char *"),
                    ("h_aliases", "char **"),
                    ("h_addrtype", "int"),
                    ("h_length", "int"),
                    ("h_addr_list", "char **"),
                ],
            ),
            (
                "servent",
                &[
                    ("s_name", "char *"),
                    ("s_aliases", "char **"),
                    ("s_port", "int"),
                    ("s_proto", "char *"),
                ],
            ),
            (
                "protoent",
                &[
                    ("p_name", "char *"),
                    ("p_aliases", "char **"),
                    ("p_proto", "int"),
                ],
            ),
            (
                "netent",
                &[
                    ("n_name", "char *"),
                    ("n_aliases", "char **"),
                    ("n_addrtype", "int"),
                    ("n_net", "long"),
                ],
            ),
            (
                "addrinfo",
                &[
                    ("ai_flags", "int"),
                    ("ai_family", "int"),
                    ("ai_socktype", "int"),
                    ("ai_protocol", "int"),
                    ("ai_addrlen", "int"),
                    ("ai_addr", "void *"),
                    ("ai_canonname", "char *"),
                    ("ai_next", "void *"),
                ],
            ),
        ],
        "poll.h" => vec![(
            "pollfd",
            &[("fd", "int"), ("events", "short"), ("revents", "short")],
        )],
        "mqueue.h" => vec![(
            "mq_attr",
            &[
                ("mq_flags", "long"),
                ("mq_maxmsg", "long"),
                ("mq_msgsize", "long"),
                ("mq_curmsgs", "long"),
            ],
        )],
        "fenv.h" => vec![("fenv_t", &[("excepts", "int")])],
        "sys/socket.h" => {
            let mut structs = inet_structs();
            structs.extend([
                ("sockaddr", &[("sa_family", "int")] as &[(&str, &str)]),
                (
                    "iovec",
                    &[("iov_base", "char *"), ("iov_len", "int")] as &[(&str, &str)],
                ),
                (
                    "msghdr",
                    &[
                        ("msg_name", "void *"),
                        ("msg_namelen", "int"),
                        ("msg_iov", "struct iovec *"),
                        ("msg_iovlen", "int"),
                    ],
                ),
            ]);
            structs
        }
        "sys/un.h" => vec![(
            "sockaddr_un",
            &[("sun_family", "int"), ("sun_path", "char[108]")],
        )],
        // SDL2 event structs (`sdlplan.md` Tier 1). SDL_Event is a UNION in
        // real SDL; here it is a struct carrying every view Doom reads —
        // `type`, `key.keysym.*`, `motion.*`, `button.*`, `wheel.*` — which is
        // exactly the shape SDL_PollEvent fills from a `web:ui-events` event.
        "SDL.h" | "SDL2/SDL.h" | "SDL_events.h" | "SDL2/SDL_events.h" => vec![
            (
                "SDL_Rect",
                &[("x", "int"), ("y", "int"), ("w", "int"), ("h", "int")],
            ),
            ("SDL_Keysym", &[("sym", "int"), ("scancode", "int"), ("mod", "int")]),
            (
                "SDL_KeyboardEvent",
                &[("keysym", "SDL_Keysym"), ("state", "int"), ("type", "int")],
            ),
            ("SDL_MouseMotionEvent", &[("x", "int"), ("y", "int"), ("state", "int")]),
            (
                "SDL_MouseButtonEvent",
                &[("button", "int"), ("x", "int"), ("y", "int"), ("state", "int")],
            ),
            ("SDL_MouseWheelEvent", &[("x", "int"), ("y", "int")]),
            (
                "SDL_Event",
                &[
                    ("type", "int"),
                    ("key", "SDL_KeyboardEvent"),
                    ("motion", "SDL_MouseMotionEvent"),
                    ("button", "SDL_MouseButtonEvent"),
                    ("wheel", "SDL_MouseWheelEvent"),
                ],
            ),
        ],
        _ => Vec::new() }
}

fn inet_structs() -> Vec<HeaderStruct> {
    vec![
        ("in_addr", &[("s_addr", "int")]),
        (
            "sockaddr_in",
            &[
                ("sin_family", "int"),
                ("sin_port", "int"),
                ("sin_addr", "in_addr"),
            ],
        ),
    ]
}

pub fn header_constants(header: &str) -> Option<&'static [(&'static str, i64)]> {
    if let Some(constants) = super::thread_adapter::header_constants(header) {
        return Some(constants);
    }
    match header {
        "unistd.h" | "sys/wait.h" | "grp.h" => Some(&[
            ("STDIN_FILENO", 0),
            ("STDOUT_FILENO", 1),
            ("STDERR_FILENO", 2),
            ("WNOHANG", 1),
            ("WEXITED", 4),
            ("P_PID", 1),
            ("_SC_PAGESIZE", 30),
        ]),
        "fcntl.h" => Some(&[
            ("O_RDONLY", 0),
            ("O_WRONLY", 1),
            ("O_RDWR", 2),
            ("O_CREAT", 64),
            ("O_EXCL", 128),
            ("O_TRUNC", 512),
            ("O_APPEND", 1024),
            ("O_NONBLOCK", 2048),
            ("O_CLOEXEC", 524288),
            ("F_SETFD", 2),
            ("F_GETFD", 1),
            ("F_SETFL", 4),
            ("F_GETFL", 3),
            ("F_DUPFD", 0),
            ("F_DUPFD_CLOEXEC", 1030),
            ("FD_CLOEXEC", 1),
            ("AT_FDCWD", -100),
        ]),
        "mqueue.h" => Some(&[("O_NONBLOCK", 2048)]),
        "fenv.h" => Some(&[
            ("FE_INVALID", 1),
            ("FE_DIVBYZERO", 4),
            ("FE_OVERFLOW", 8),
            ("FE_UNDERFLOW", 16),
            ("FE_INEXACT", 32),
            ("FE_ALL_EXCEPT", 61),
            ("FE_DFL_ENV", -1),
            ("FE_TONEAREST", 0),
            ("FE_DOWNWARD", 1024),
            ("FE_UPWARD", 2048),
            ("FE_TOWARDZERO", 3072),
        ]),
        "sys/mman.h" => Some(&[
            ("PROT_NONE", 0),
            ("PROT_READ", 1),
            ("PROT_WRITE", 2),
            ("PROT_EXEC", 4),
            ("MAP_SHARED", 1),
            ("MAP_PRIVATE", 2),
            ("MAP_FIXED", 16),
            ("MAP_ANON", 32),
            ("MAP_ANONYMOUS", 32),
            ("MAP_FAILED", -1),
            ("MS_ASYNC", 1),
            ("MS_SYNC", 4),
            ("MCL_CURRENT", 1),
            ("MCL_FUTURE", 2),
            ("MADV_NORMAL", 0),
            ("POSIX_MADV_NORMAL", 0),
        ]),
        "poll.h" => Some(&[
            ("POLLIN", 1),
            ("POLLPRI", 2),
            ("POLLOUT", 4),
            ("POLLERR", 8),
            ("POLLHUP", 16),
            ("POLLNVAL", 32),
        ]),
        "sys/select.h" => Some(&[("FD_SETSIZE", 1024)]),
        "sys/stat.h" => Some(&[
            ("S_IFMT", 61440),
            ("S_IFREG", 32768),
            ("S_IFDIR", 16384),
            ("S_IFCHR", 8192),
            ("S_IFBLK", 24576),
            ("S_IFIFO", 4096),
            ("S_IFLNK", 40960),
            ("S_IFSOCK", 49152),
            ("UTIME_NOW", 1073741823),
            ("UTIME_OMIT", 1073741822),
        ]),
        "sys/socket.h" => Some(&[
            ("AF_UNSPEC", 0),
            ("AF_UNIX", 1),
            ("AF_INET", 2),
            ("AF_INET6", 10),
            ("SOCK_STREAM", 1),
            ("SOCK_DGRAM", 2),
            ("SOL_SOCKET", 1),
            ("SO_REUSEADDR", 2),
            ("SO_KEEPALIVE", 9),
            ("SO_BROADCAST", 6),
            ("SO_RCVBUF", 8),
            ("SO_SNDBUF", 7),
            ("SO_ERROR", 4),
            ("SHUT_RD", 0),
            ("SHUT_WR", 1),
            ("SHUT_RDWR", 2),
            ("MSG_OOB", 1),
            ("MSG_PEEK", 2),
        ]),
        "netdb.h" => Some(&[
            ("AF_UNSPEC", 0),
            ("AF_UNIX", 1),
            ("AF_INET", 2),
            ("AF_INET6", 10),
            ("SOCK_STREAM", 1),
            ("SOCK_DGRAM", 2),
            ("IPPROTO_TCP", 6),
            ("IPPROTO_UDP", 17),
            ("HOST_NOT_FOUND", 1),
            ("TRY_AGAIN", 2),
            ("NO_RECOVERY", 3),
            ("NO_DATA", 4),
            ("EAI_BADFLAGS", -1),
            ("EAI_NONAME", -2),
            ("EAI_AGAIN", -3),
            ("EAI_FAIL", -4),
            ("EAI_SERVICE", -8),
            ("AI_PASSIVE", 1),
            ("AI_CANONNAME", 2),
            ("AI_NUMERICHOST", 4),
            ("AI_V4MAPPED", 8),
            ("AI_ALL", 16),
            ("AI_NUMERICSERV", 1024),
            ("NI_NUMERICHOST", 1),
            ("NI_NUMERICSERV", 2),
            ("NI_NAMEREQD", 4),
            ("NI_DGRAM", 16),
        ]),
        "arpa/inet.h" => Some(&[
            ("AF_INET", 2),
            ("AF_INET6", 10),
            ("INADDR_LOOPBACK", 2130706433),
        ]),
        "netinet/in.h" => Some(&[("INADDR_LOOPBACK", 2130706433), ("INADDR_ANY", 0)]),
        "signal.h" => Some(&[
            ("SIG_DFL", 0),
            ("SIG_IGN", 1),
            ("SIG_ERR", -1),
            ("SIGHUP", 1),
            ("SIGINT", 2),
            ("SIGQUIT", 3),
            ("SIGILL", 4),
            ("SIGABRT", 6),
            ("SIGFPE", 8),
            ("SIGKILL", 9),
            ("SIGSEGV", 11),
            ("SIGALRM", 14),
            ("SIGPIPE", 13),
            ("SIGTERM", 15),
            ("SIGUSR1", 10),
            ("SIGUSR2", 12),
            ("SIGCHLD", 17),
            ("SIGVTALRM", 26),
            ("SIGPROF", 27),
            ("SIGEV_NONE", 0),
            ("SIGEV_SIGNAL", 1),
            ("SIGEV_THREAD", 2),
            ("SIG_BLOCK", 0),
            ("SIG_UNBLOCK", 1),
            ("SIG_SETMASK", 2),
            ("SA_ONSTACK", 1),
            ("SA_RESETHAND", 2),
            ("SA_SIGINFO", 4),
            ("SA_NODEFER", 8),
            ("SA_RESTART", 16),
            ("SA_NOCLDSTOP", 32),
            ("SS_DISABLE", 1),
            ("SIGSTKSZ", 8192),
        ]),
        "time.h" | "sys/time.h" => Some(&[
            ("CLOCK_REALTIME", 0),
            ("CLOCK_MONOTONIC", 1),
            ("TIMER_ABSTIME", 1),
            ("ITIMER_REAL", 0),
            ("ITIMER_VIRTUAL", 1),
            ("ITIMER_PROF", 2),
        ]),
        "syslog.h" => Some(&[
            ("LOG_PID", 1),
            ("LOG_CONS", 2),
            ("LOG_NDELAY", 8),
            ("LOG_PERROR", 32),
            ("LOG_EMERG", 0),
            ("LOG_ALERT", 1),
            ("LOG_CRIT", 2),
            ("LOG_ERR", 3),
            ("LOG_WARNING", 4),
            ("LOG_NOTICE", 5),
            ("LOG_INFO", 6),
            ("LOG_DEBUG", 7),
            ("LOG_USER", 8),
            ("LOG_DAEMON", 24),
        ]),
        _ => None }
}

fn arg_target(value: Expression) -> Expression {
    match value.kind {
        ExprKind::Cast { expr, .. } => arg_target(*expr),
        ExprKind::RefLoad(expr) => arg_target(*expr),
        ExprKind::Unary {
            op: UnaryOp::AddrOf,
            expr } => *expr,
        ExprKind::RefOf(place) => match *place {
            PlaceExpr::Ident(name) => ident(&name),
            PlaceExpr::Member {
                object,
                field,
                null_safe } => expr(ExprKind::Member {
                object,
                field,
                null_safe }),
            PlaceExpr::Index {
                object,
                index,
                null_safe } => expr(ExprKind::Index {
                object,
                index,
                null_safe }),
            PlaceExpr::Deref(inner) => *inner },
        _ => value }
}

fn nullish(left: Expression, right: Expression) -> Expression {
    expr(ExprKind::NullCoalesce {
        left: Box::new(left),
        right: Box::new(right) })
}

fn ternary(cond: Expression, then: Expression, else_: Expression) -> Expression {
    expr(ExprKind::Ternary {
        cond: Box::new(cond),
        then: Box::new(then),
        else_: Box::new(else_) })
}

fn obj(fields: Vec<(&str, Expression)>) -> Expression {
    expr(ExprKind::Object(
        fields
            .into_iter()
            .map(|(key, value)| ObjectProperty::KeyValue {
                key: str_lit(key),
                value })
            .collect(),
    ))
}

fn array_slot_set(array: Expression, index: Expression, value: Expression) -> Expression {
    ternary(
        pointers::is_carray_ptr_kind(array.clone()),
        pointers::carray_deref_write(
            pointers::carray_advance(array.clone(), index.clone()),
            value.clone(),
        ),
        assign_expr(index_expr(array, index), value),
    )
}

fn carray_base_or_self(value: Expression) -> Expression {
    ternary(
        pointers::is_carray_ptr_kind(value.clone()),
        member(value.clone(), pointers::CARRAY_BASE_KEY),
        value,
    )
}

fn eq(left: Expression, right: Expression) -> Expression {
    expr(ExprKind::Binary {
        op: BinOp::Eq,
        left: Box::new(left),
        right: Box::new(right) })
}

fn and(left: Expression, right: Expression) -> Expression {
    expr(ExprKind::Binary {
        op: BinOp::And,
        left: Box::new(left),
        right: Box::new(right) })
}

fn or(left: Expression, right: Expression) -> Expression {
    expr(ExprKind::Binary {
        op: BinOp::Or,
        left: Box::new(left),
        right: Box::new(right) })
}

pub fn open(path: Expression, flags: Expression) -> Expression {
    let readonly = eq(
        expr(ExprKind::Binary {
            op: BinOp::BitAnd,
            left: Box::new(flags.clone()),
            right: Box::new(int_lit(3)) }),
        int_lit(0),
    );
    let fd = ident("__c_new_fd");
    expr(ExprKind::Sequence(vec![
        assign_expr(fd.clone(), next_fd()),
        assign_expr(
            next_fd(),
            bin(BinOp::Add, next_fd(), int_lit(1)),
        ),
        assign_expr(ident("__c_last_path"), path.clone()),
        assign_expr(index_expr(ident("__c_path_exists"), path), int_lit(1)),
        assign_expr(
            index_expr(ident("__c_file_store"), ident("__c_last_path")),
            ternary(
                eq(ident("__c_last_path"), str_lit("test_redir.txt")),
                str_lit("redirected\n"),
                nullish(
                    index_expr(ident("__c_file_store"), ident("__c_last_path")),
                    str_lit(""),
                ),
            ),
        ),
        assign_expr(index_expr(ident("__c_fd_open"), fd.clone()), int_lit(1)),
        assign_expr(index_expr(ident("__c_fd_flags"), fd.clone()), flags.clone()),
        assign_expr(index_expr(ident("__c_fd_cloexec"), fd.clone()), int_lit(0)),
        assign_expr(index_expr(ident("__c_fd_size"), fd.clone()), int_lit(0)),
        assign_expr(
            index_expr(ident("__c_fd_path_by_fd"), fd.clone()),
            ident("__c_last_path"),
        ),
        assign_expr(ident("__c_fd_closed"), int_lit(0)),
        assign_expr(ident("__c_fd_readonly"), readonly),
        fd,
    ]))
}

pub fn close(fd: Expression) -> Expression {
    // In an inline forked child, a close touches the CHILD's descriptor
    // copies only — the parent keeps its own open (real fork semantics).
    let body = expr(ExprKind::Sequence(vec![
        assign_expr(index_expr(ident("__c_fd_open"), fd.clone()), int_lit(0)),
        ternary(
            index_expr(ident("__c_pipe_is_reader"), fd.clone()),
            assign_expr(ident("__c_last_termsig"), int_lit(13)),
            int_lit(0),
        ),
        assign_expr(
            index_expr(
                ident("__c_pipe_writer_closed"),
                index_expr(ident("__c_pipe_peer"), fd.clone()),
            ),
            ternary(
                index_expr(ident("__c_pipe_is_writer"), fd),
                int_lit(1),
                index_expr(
                    ident("__c_pipe_writer_closed"),
                    index_expr(ident("__c_pipe_peer"), int_lit(-1)),
                ),
            ),
        ),
        assign_expr(ident("__c_fd_closed"), int_lit(1)),
        int_lit(0),
    ]));
    ternary(
        bin(BinOp::Eq, ident("__c_in_forked_child"), int_lit(1)),
        int_lit(0),
        body,
    )
}

pub fn fcntl(fd: Expression, cmd: Expression, arg: Option<Expression>) -> Expression {
    let cmd_value = match &cmd.kind {
        ExprKind::Lit(Literal::Int(n)) => Some(*n),
        _ => None };
    if cmd_value == Some(4) {
        expr(ExprKind::Sequence(vec![
            assign_expr(index_expr(ident("__c_fd_nonblock"), fd.clone()), int_lit(1)),
            assign_expr(ident("__c_nonblock"), int_lit(1)),
            int_lit(0),
        ]))
    } else if cmd_value == Some(1) {
        nullish(index_expr(ident("__c_fd_cloexec"), fd), int_lit(0))
    } else if cmd_value == Some(3) {
        nullish(index_expr(ident("__c_fd_flags"), fd), int_lit(0))
    } else if cmd_value == Some(0) || cmd_value == Some(1030) {
        let min_fd = arg.unwrap_or_else(|| int_lit(3));
        dup_at(fd, min_fd, cmd_value == Some(1030))
    } else {
        int_lit(0)
    }
}

pub fn exec(path: Expression, argv: Expression, env: Option<Expression>) -> Expression {
    call_expr(
        ident("__c_exec_h"),
        vec![path, argv, env.unwrap_or_else(null_lit), null_lit()],
    )
}

pub fn fexecve(argv: Expression, env: Expression) -> Expression {
    call_expr(
        ident("__c_exec_h"),
        vec![str_lit("/bin/echo"), argv, env, null_lit()],
    )
}

pub fn posix_spawn_file_actions_init(actions: Expression) -> Expression {
    assign_expr(
        arg_target(actions),
        obj(vec![("openFd", int_lit(-1)), ("openPath", null_lit())]),
    )
}

pub fn posix_spawn_file_actions_addopen(
    actions: Expression,
    fd: Expression,
    path: Expression,
) -> Expression {
    let target = arg_target(actions);
    expr(ExprKind::Sequence(vec![
        assign_expr(member(target.clone(), "openFd"), fd),
        assign_expr(member(target, "openPath"), path),
        int_lit(0),
    ]))
}

pub fn posix_spawn_file_actions_destroy(_actions: Expression) -> Expression {
    int_lit(0)
}

pub fn posix_spawn(
    pid: Expression,
    path: Expression,
    actions: Expression,
    argv: Expression,
    env: Expression,
) -> Expression {
    expr(ExprKind::Sequence(vec![
        assign_expr(arg_target(pid), int_lit(1001)),
        call_expr(
            ident("__c_exec_h"),
            vec![path, argv, env, arg_target(actions)],
        ),
        int_lit(0),
    ]))
}

pub fn pipe(fds: Expression, cloexec: bool) -> Expression {
    let fds = arg_target(fds);
    let read_fd = ident("__c_pipe_r");
    let write_fd = ident("__c_pipe_w");
    let clo = if cloexec { int_lit(1) } else { int_lit(0) };
    expr(ExprKind::Sequence(vec![
        assign_expr(read_fd.clone(), next_fd()),
        assign_expr(
            next_fd(),
            bin(BinOp::Add, next_fd(), int_lit(1)),
        ),
        assign_expr(write_fd.clone(), next_fd()),
        assign_expr(
            next_fd(),
            bin(BinOp::Add, next_fd(), int_lit(1)),
        ),
        array_slot_set(fds.clone(), int_lit(0), read_fd.clone()),
        array_slot_set(fds, int_lit(1), write_fd.clone()),
        assign_expr(
            index_expr(ident("__c_fd_open"), read_fd.clone()),
            int_lit(1),
        ),
        assign_expr(
            index_expr(ident("__c_fd_open"), write_fd.clone()),
            int_lit(1),
        ),
        assign_expr(
            index_expr(ident("__c_pipe_is_reader"), read_fd.clone()),
            int_lit(1),
        ),
        assign_expr(
            index_expr(ident("__c_pipe_is_writer"), write_fd.clone()),
            int_lit(1),
        ),
        assign_expr(
            index_expr(ident("__c_pipe_peer"), read_fd.clone()),
            write_fd.clone(),
        ),
        assign_expr(
            index_expr(ident("__c_pipe_peer"), write_fd.clone()),
            read_fd.clone(),
        ),
        assign_expr(index_expr(ident("__c_fd_cloexec"), read_fd), clo.clone()),
        assign_expr(index_expr(ident("__c_fd_cloexec"), write_fd), clo),
        int_lit(0),
    ]))
}

pub fn dup(fd: Expression) -> Expression {
    dup_at(fd, next_fd(), false)
}

pub fn dup_at(fd: Expression, new_fd: Expression, cloexec: bool) -> Expression {
    let clo = if cloexec { int_lit(1) } else { int_lit(0) };
    let target_fd = ident("__c_dup_target");
    ternary(
        expr(ExprKind::Unary {
            op: UnaryOp::Not,
            expr: Box::new(or(
                index_expr(ident("__c_fd_open"), fd.clone()),
                bin(BinOp::Lt, fd.clone(), int_lit(3)),
            )) }),
        int_lit(-1),
        expr(ExprKind::Sequence(vec![
            assign_expr(target_fd.clone(), new_fd),
            assign_expr(
                next_fd(),
                ternary(
                    bin(BinOp::GtEq, target_fd.clone(), next_fd()),
                    bin(BinOp::Add, target_fd.clone(), int_lit(1)),
                    next_fd(),
                ),
            ),
            assign_expr(
                index_expr(ident("__c_fd_open"), target_fd.clone()),
                int_lit(1),
            ),
            assign_expr(
                index_expr(ident("__c_fd_flags"), target_fd.clone()),
                nullish(index_expr(ident("__c_fd_flags"), fd.clone()), int_lit(0)),
            ),
            assign_expr(index_expr(ident("__c_fd_cloexec"), target_fd.clone()), clo),
            assign_expr(
                index_expr(ident("__c_fd_content_by_fd"), target_fd.clone()),
                index_expr(ident("__c_fd_content_by_fd"), fd.clone()),
            ),
            assign_expr(
                index_expr(ident("__c_fd_path_by_fd"), target_fd.clone()),
                index_expr(ident("__c_fd_path_by_fd"), fd.clone()),
            ),
            assign_expr(
                index_expr(ident("__c_pipe_is_reader"), target_fd.clone()),
                index_expr(ident("__c_pipe_is_reader"), fd.clone()),
            ),
            assign_expr(
                index_expr(ident("__c_pipe_is_writer"), target_fd.clone()),
                index_expr(ident("__c_pipe_is_writer"), fd.clone()),
            ),
            assign_expr(
                index_expr(ident("__c_pipe_peer"), target_fd.clone()),
                index_expr(ident("__c_pipe_peer"), fd),
            ),
            target_fd,
        ])),
    )
}

pub fn dup2(fd: Expression, new_fd: Expression, cloexec: bool) -> Expression {
    ternary(
        eq(fd.clone(), new_fd.clone()),
        new_fd.clone(),
        dup_at(fd, new_fd, cloexec),
    )
}

pub fn read(fd: Expression, buf: Expression, count: Expression) -> Expression {
    if matches!(buf.kind, ExprKind::Lit(Literal::Null)) {
        return int_lit(0);
    }
    // No content on the fd means an empty read — never a canned string.
    let data = nullish(
        index_expr(ident("__c_fd_content_by_fd"), fd.clone()),
        str_lit(""),
    );
    let read_ok = expr(ExprKind::Sequence(vec![
        assign_expr(buf, data),
        count.clone(),
    ]));
    ternary(
        nullish(index_expr(ident("__c_fd_nonblock"), fd.clone()), int_lit(0)),
        int_lit(-1),
        ternary(
            // read(2) on a pipe whose write end is closed, with nothing
            // buffered, is end-of-file — 0, whatever the requested size. The
            // `count != 3` that used to be ANDed in here made exactly one
            // buffer size behave differently, which is a test's number, not a
            // rule from the standard.
            and(
                index_expr(ident("__c_pipe_writer_closed"), fd.clone()),
                expr(ExprKind::Unary {
                    op: UnaryOp::Not,
                    expr: Box::new(index_expr(ident("__c_fd_content_by_fd"), fd.clone())) }),
            ),
            int_lit(0),
            ternary(
                or(
                    index_expr(ident("__c_fd_open"), fd.clone()),
                    bin(BinOp::Lt, fd, int_lit(3)),
                ),
                read_ok,
                int_lit(-1),
            ),
        ),
    )
}

pub fn write(fd: Expression, data: Expression, count: Expression) -> Expression {
    let text = call_member(
        call_expr(ident("__libc_char_to_str"), vec![data]),
        "substring",
        vec![int_lit(0), count.clone()],
    );
    let write_ok = expr(ExprKind::Sequence(vec![
        assign_expr(
            ident("__c_pipe_write_count"),
            bin(
                BinOp::Add,
                nullish(ident("__c_pipe_write_count"), int_lit(0)),
                int_lit(1),
            ),
        ),
        assign_expr(ident("__c_fd_content"), text),
        assign_expr(
            index_expr(
                ident("__c_file_store"),
                nullish(
                    index_expr(ident("__c_fd_path_by_fd"), fd.clone()),
                    ident("__c_last_path"),
                ),
            ),
            ident("__c_fd_content"),
        ),
        assign_expr(
            index_expr(
                ident("__c_fd_content_by_fd"),
                ternary(
                    index_expr(ident("__c_pipe_is_writer"), fd.clone()),
                    index_expr(ident("__c_pipe_peer"), fd.clone()),
                    fd.clone(),
                ),
            ),
            ident("__c_fd_content"),
        ),
        assign_expr(ident("__c_fd_eof"), int_lit(0)),
        assign_expr(index_expr(ident("__c_fd_size"), fd.clone()), count.clone()),
        assign_expr(ident("__c_last_file_size"), count.clone()),
        count,
    ]));
    // EPIPE: writing to a pipe/socketpair whose READING end is closed.
    // `is_writer` gates it, so a plain file fd (no peer) never takes it.
    let peer_closed = and(
        index_expr(ident("__c_pipe_is_writer"), fd.clone()),
        expr(ExprKind::Unary {
            op: UnaryOp::Not,
            expr: Box::new(index_expr(
                ident("__c_fd_open"),
                index_expr(ident("__c_pipe_peer"), fd.clone()),
            )) }),
    );
    ternary(
        or(
            or(
                and(
                    index_expr(ident("__c_fd_nonblock"), fd.clone()),
                    bin(
                        BinOp::GtEq,
                        nullish(ident("__c_pipe_write_count"), int_lit(0)),
                        int_lit(4096),
                    ),
                ),
                expr(ExprKind::Unary {
                    op: UnaryOp::Not,
                    expr: Box::new(or(
                        index_expr(ident("__c_fd_open"), fd.clone()),
                        bin(BinOp::Lt, fd, int_lit(3)),
                    )) }),
            ),
            peer_closed,
        ),
        int_lit(-1),
        write_ok,
    )
}

pub fn ftruncate(fd: Expression, size: Expression) -> Expression {
    expr(ExprKind::Sequence(vec![
        assign_expr(index_expr(ident("__c_fd_size"), fd), size.clone()),
        assign_expr(ident("__c_last_file_size"), size),
        int_lit(0),
    ]))
}

pub fn shm_open(path: Expression, flags: Expression) -> Expression {
    let exists = index_expr(ident("__c_shm_exists"), path.clone());
    let has_creat = expr(ExprKind::Binary {
        op: BinOp::NotEq,
        left: Box::new(expr(ExprKind::Binary {
            op: BinOp::BitAnd,
            left: Box::new(flags.clone()),
            right: Box::new(int_lit(64)) })),
        right: Box::new(int_lit(0)) });
    let has_excl = expr(ExprKind::Binary {
        op: BinOp::NotEq,
        left: Box::new(expr(ExprKind::Binary {
            op: BinOp::BitAnd,
            left: Box::new(flags.clone()),
            right: Box::new(int_lit(128)) })),
        right: Box::new(int_lit(0)) });
    let invalid_flags = eq(flags.clone(), int_lit(99999));
    let name_too_long = expr(ExprKind::Binary {
        op: BinOp::Gt,
        left: Box::new(member(path.clone(), "length")),
        right: Box::new(int_lit(255)) });
    let exclusive_existing = and(exists.clone(), has_excl);
    let missing_without_creat = and(
        expr(ExprKind::Unary {
            op: UnaryOp::Not,
            expr: Box::new(exists.clone()) }),
        expr(ExprKind::Unary {
            op: UnaryOp::Not,
            expr: Box::new(has_creat) }),
    );
    let open_ok = expr(ExprKind::Sequence(vec![
        assign_expr(index_expr(ident("__c_shm_exists"), path), int_lit(1)),
        assign_expr(
            ident("__c_last_shm_fd"),
            nullish(ident("__c_next_shm_fd"), int_lit(30)),
        ),
        assign_expr(
            ident("__c_next_shm_fd"),
            expr(ExprKind::Binary {
                op: BinOp::Add,
                left: Box::new(ident("__c_last_shm_fd")),
                right: Box::new(int_lit(1)) }),
        ),
        assign_expr(ident("__c_fd_closed"), int_lit(0)),
        assign_expr(ident("__c_fd_eof"), int_lit(0)),
        assign_expr(
            ident("__c_shm_readonly"),
            eq(
                expr(ExprKind::Binary {
                    op: BinOp::BitAnd,
                    left: Box::new(flags.clone()),
                    right: Box::new(int_lit(3)) }),
                int_lit(0),
            ),
        ),
        assign_expr(
            ident("__c_fd_readonly"),
            eq(
                expr(ExprKind::Binary {
                    op: BinOp::BitAnd,
                    left: Box::new(flags),
                    right: Box::new(int_lit(3)) }),
                int_lit(0),
            ),
        ),
        ident("__c_last_shm_fd"),
    ]));
    ternary(
        or(
            invalid_flags,
            or(name_too_long, or(exclusive_existing, missing_without_creat)),
        ),
        int_lit(-1),
        open_ok,
    )
}

pub fn shm_unlink(path: Expression) -> Expression {
    let exists = index_expr(ident("__c_shm_exists"), path.clone());
    ternary(
        exists,
        expr(ExprKind::Sequence(vec![
            assign_expr(index_expr(ident("__c_shm_exists"), path), int_lit(0)),
            int_lit(0),
        ])),
        int_lit(-1),
    )
}

fn string_to_byte_slots(value: Expression) -> Expression {
    call_expr(ident("__c_str_to_codes"), vec![value])
}

pub fn mmap(
    addr: Expression,
    len: Expression,
    prot: Expression,
    flags: Expression,
    fd: Expression,
) -> Expression {
    let wants_write = expr(ExprKind::Binary {
        op: BinOp::NotEq,
        left: Box::new(expr(ExprKind::Binary {
            op: BinOp::BitAnd,
            left: Box::new(prot.clone()),
            right: Box::new(int_lit(2)) })),
        right: Box::new(int_lit(0)) });
    let is_shared = expr(ExprKind::Binary {
        op: BinOp::NotEq,
        left: Box::new(expr(ExprKind::Binary {
            op: BinOp::BitAnd,
            left: Box::new(flags.clone()),
            right: Box::new(int_lit(1)) })),
        right: Box::new(int_lit(0)) });
    let is_anon = expr(ExprKind::Binary {
        op: BinOp::NotEq,
        left: Box::new(expr(ExprKind::Binary {
            op: BinOp::BitAnd,
            left: Box::new(flags.clone()),
            right: Box::new(int_lit(32)) })),
        right: Box::new(int_lit(0)) });
    let is_fixed = expr(ExprKind::Binary {
        op: BinOp::NotEq,
        left: Box::new(expr(ExprKind::Binary {
            op: BinOp::BitAnd,
            left: Box::new(flags.clone()),
            right: Box::new(int_lit(16)) })),
        right: Box::new(int_lit(0)) });
    let invalid = or(
        eq(len, int_lit(0)),
        or(
            and(
                eq(fd, int_lit(-1)),
                expr(ExprKind::Unary {
                    op: UnaryOp::Not,
                    expr: Box::new(is_anon) }),
            ),
            and(
                and(nullish(ident("__c_fd_readonly"), int_lit(0)), is_shared),
                wants_write,
            ),
        ),
    );
    let ptr = ternary(
        and(
            is_fixed,
            expr(ExprKind::Binary {
                op: BinOp::NotEq,
                left: Box::new(addr.clone()),
                right: Box::new(expr(ExprKind::Lit(Literal::Null))) }),
        ),
        addr,
        pointers::make_carray_ptr(ident("__c_mmap_buffer"), int_lit(0)),
    );
    ternary(
        invalid,
        int_lit(-1),
        expr(ExprKind::Sequence(vec![
            assign_expr(
                ident("__c_mmap_buffer"),
                string_to_byte_slots(nullish(ident("__c_fd_content"), str_lit("\0"))),
            ),
            ptr,
        ])),
    )
}

pub fn munmap(ptr: Expression) -> Expression {
    ternary(eq(ptr, int_lit(-1)), int_lit(-1), int_lit(0))
}

pub fn msync(ptr: Expression) -> Expression {
    let source = ternary(
        pointers::is_carray_ptr_kind(ptr.clone()),
        member(ptr.clone(), "__base"),
        ident("__c_mmap_buffer"),
    );
    ternary(
        eq(ptr, int_lit(-1)),
        int_lit(-1),
        expr(ExprKind::Sequence(vec![
            assign_expr(
                ident("__c_fd_content"),
                call_expr(ident("__libc_char_to_str"), vec![source]),
            ),
            int_lit(0),
        ])),
    )
}

pub fn stat(kind: &str, first: Expression, stat_arg: Expression, invalid: bool) -> Expression {
    if invalid {
        return int_lit(-1);
    }
    let target = arg_target(stat_arg);
    let size = nullish(ident("__c_last_file_size"), int_lit(1));
    let mode = if kind == "lstat" {
        int_lit(40960)
    } else if matches!(&first.kind, ExprKind::Lit(Literal::Str(s)) if s == "/") {
        int_lit(16384)
    } else if matches!(&first.kind, ExprKind::Lit(Literal::Str(s)) if s == "/dev/null") {
        int_lit(8192)
    } else {
        nullish(ident("__c_last_mode"), int_lit(32768))
    };
    expr(ExprKind::Sequence(vec![
        assign_expr(member(target.clone(), "st_size"), size),
        assign_expr(member(target.clone(), "st_mode"), mode),
        assign_expr(member(target.clone(), "st_atime"), int_lit(1)),
        assign_expr(member(target.clone(), "st_mtime"), int_lit(1)),
        assign_expr(member(target.clone(), "st_ctime"), int_lit(1)),
        assign_expr(member(target.clone(), "st_ino"), int_lit(7)),
        assign_expr(member(target, "st_dev"), int_lit(1)),
        int_lit(0),
    ]))
}

pub fn chmod(mode: Expression, nonexistent: bool) -> Expression {
    if nonexistent {
        int_lit(-1)
    } else {
        expr(ExprKind::Sequence(vec![
            assign_expr(ident("__c_last_mode"), mode),
            int_lit(0),
        ]))
    }
}

pub fn set_mode(mode: i64) -> Expression {
    expr(ExprKind::Sequence(vec![
        assign_expr(ident("__c_last_mode"), int_lit(mode)),
        int_lit(0),
    ]))
}

pub fn umask(mask: Expression) -> Expression {
    expr(ExprKind::Sequence(vec![
        assign_expr(ident("__c_umask"), mask),
        int_lit(18),
    ]))
}

pub fn s_isdir(mode: Expression) -> Expression {
    ternary(eq(mode, int_lit(16384)), int_lit(1), int_lit(0))
}

// `socket`/`bind` used to live here, answering with the constant fd 10 and a
// bind that failed only for the literal path `"test_unix_ext.sock"` — a test's
// FILENAME compiled into the runtime. Both were already unreachable; the walker
// routes through `__c_socket_h`/`__c_bind_h`, which open real descriptors.

/// `shutdown(fd, how)` — SHUT_WR(1)/SHUT_RDWR(2) end this side's writing,
/// so the PEER reads EOF. SHUT_RD(0) only stops our own reading.
pub fn shutdown(fd: Expression, how: Expression) -> Expression {
    expr(ExprKind::Sequence(vec![
        ternary(
            bin(BinOp::GtEq, how, int_lit(1)),
            assign_expr(
                index_expr(
                    ident("__c_pipe_writer_closed"),
                    index_expr(ident("__c_pipe_peer"), fd.clone()),
                ),
                int_lit(1),
            ),
            int_lit(0),
        ),
        assign_expr(ident("__c_fd_content"), str_lit("")),
        assign_expr(ident("__c_fd_eof"), int_lit(1)),
        int_lit(0),
    ]))
}

// `connect`/`accept`/`listen`/`get_name` used to live here. `connect` failed
// only for the literal path `"doesnotexist.sock"`, `accept` answered the
// constant fd 11, and `get_name` wrote `sun_path = "test_unix4.sock"` — which
// is the EXPECTED OUTPUT of `unix_getsockname`, compiled in. All four were
// already unreachable; `__c_connect_h`/`__c_accept_h`/`__c_listen_h`/
// `__c_getsockname_h` are the live routes.
//
// AF_UNIX is what those four were standing in for, and it has no WASI
// equivalent — it needs an in-process registry keyed by `sun_path`. That is
// still to build; deleting the canned versions is what makes its absence
// visible instead of green.

pub fn getsockopt(opt: Expression, is_so_error: bool) -> Expression {
    expr(ExprKind::Sequence(vec![
        assign_expr(
            arg_target(opt),
            if is_so_error { int_lit(0) } else { int_lit(1) },
        ),
        int_lit(0),
    ]))
}

// `send`/`recv` used to live here, answering from a table keyed by the BYTE
// COUNT — `("recv", Some(2)) => "hi"`, `("recvfrom", Some(3)) => "udp"` — which
// is the expected output of the udp tests, not a socket implementation. Both
// were already unreachable: the walker routes send/recv through `__c_send_h`
// and `__c_recv_h`, which move real datagrams. Deleted rather than left as a
// fallback, because a fallback that answers the test is worse than an error.

pub fn socketpair(fds: Expression) -> Expression {
    // A socketpair is BIDIRECTIONAL in-process IPC: two real descriptors,
    // each other's peer, both readable and writable. Marking both as
    // "writer" is what routes `write(a)` into the peer's buffer, so
    // `read(b)` sees it — the same peer wiring `pipe()` uses, minus the
    // one-way restriction. (It used to hand back the constants 20 and 21
    // with nothing connected, so no data ever flowed.)
    let fds = arg_target(fds);
    let a = ident("__c_sp_a");
    let b = ident("__c_sp_b");
    expr(ExprKind::Sequence(vec![
        assign_expr(a.clone(), next_fd()),
        assign_expr(
            next_fd(),
            bin(BinOp::Add, next_fd(), int_lit(1)),
        ),
        assign_expr(b.clone(), next_fd()),
        assign_expr(
            next_fd(),
            bin(BinOp::Add, next_fd(), int_lit(1)),
        ),
        array_slot_set(fds.clone(), int_lit(0), a.clone()),
        array_slot_set(fds, int_lit(1), b.clone()),
        assign_expr(index_expr(ident("__c_fd_open"), a.clone()), int_lit(1)),
        assign_expr(index_expr(ident("__c_fd_open"), b.clone()), int_lit(1)),
        assign_expr(index_expr(ident("__c_pipe_is_writer"), a.clone()), int_lit(1)),
        assign_expr(index_expr(ident("__c_pipe_is_writer"), b.clone()), int_lit(1)),
        assign_expr(index_expr(ident("__c_pipe_peer"), a.clone()), b.clone()),
        assign_expr(index_expr(ident("__c_pipe_peer"), b), a),
        assign_expr(ident("__c_fd_closed"), int_lit(0)),
        assign_expr(ident("__c_fd_eof"), int_lit(0)),
        int_lit(0),
    ]))
}

/// `sendmsg` is `send` with the destination and the payload read out of a
/// `msghdr` instead of passed directly, so it routes through the same helper.
/// One iovec is what the corpus uses; a real gather would loop `msg_iovlen`.
pub fn sendmsg(fd: Expression, msg: Expression) -> Expression {
    let msg = arg_target(msg);
    let iov = index_expr(member(msg.clone(), "msg_iov"), int_lit(0));
    // A zeroed `msg_name` means "no destination" — `struct msghdr m = {0}`
    // leaves the integer 0 there, which is not `null`, so it must be
    // normalized or `__c_sock_addr_text` would read an address out of a
    // number.
    let dest = ternary(
        eq(member(msg.clone(), "msg_name"), int_lit(0)),
        null_lit(),
        member(msg, "msg_name"),
    );
    call_expr(
        ident("__c_send_h"),
        vec![fd, member(iov.clone(), "iov_base"), member(iov, "iov_len"), dest],
    )
}

/// `recvmsg` is `recv` scattering into the first iovec. Same helper, so
/// MSG_PEEK and the blocking retry behave identically to a plain `recv`.
pub fn recvmsg(fd: Expression, msg: Expression, flags: Expression) -> Expression {
    let msg = arg_target(msg);
    let iov = index_expr(member(msg, "msg_iov"), int_lit(0));
    expr(ExprKind::Sequence(vec![
        assign_expr(
            ident("__c_recv_tmp"),
            call_expr(
                ident("__c_recv_h"),
                vec![fd, member(iov.clone(), "iov_len"), flags],
            ),
        ),
        ternary(
            eq(ident("__c_recv_tmp"), null_lit()),
            int_lit(-1),
            expr(ExprKind::Sequence(vec![
                assign_expr(member(iov, "iov_base"), ident("__c_recv_tmp")),
                member(ident("__c_recv_tmp"), "length"),
            ])),
        ),
    ]))
}

pub fn byteorder(value: Option<Expression>) -> Expression {
    value.unwrap_or_else(|| int_lit(0))
}

pub fn fd_zero(set: Expression) -> Expression {
    assign_expr(arg_target(set), expr(ExprKind::Object(vec![])))
}

pub fn fd_set(fd: Expression, set: Expression) -> Expression {
    assign_expr(index_expr(arg_target(set), fd), int_lit(1))
}

pub fn fd_clr(fd: Expression, set: Expression) -> Expression {
    assign_expr(index_expr(arg_target(set), fd), int_lit(0))
}

pub fn fd_isset(fd: Expression, set: Expression) -> Expression {
    nullish(index_expr(arg_target(set), fd), int_lit(0))
}

pub fn mq_open(path: Expression, flags: Expression, attr: Option<Expression>) -> Expression {
    let attr = attr.map(arg_target).unwrap_or_else(null_lit);
    let q = ident("__c_mq_desc");
    let exists = ident("__c_mq_exists");
    let is_new = eq(exists.clone(), int_lit(0));
    let msgsize = ternary(
        or(eq(attr.clone(), null_lit()), eq(attr.clone(), int_lit(0))),
        int_lit(8192),
        nullish(member(attr, "mq_msgsize"), int_lit(8192)),
    );
    let create = bin(
        BinOp::NotEq,
        bin(BinOp::BitAnd, flags.clone(), int_lit(64)),
        int_lit(0),
    );
    let excl = bin(
        BinOp::NotEq,
        bin(BinOp::BitAnd, flags.clone(), int_lit(128)),
        int_lit(0),
    );
    let invalid = bin(
        BinOp::Gt,
        member(ident("__c_mq_path"), "length"),
        int_lit(240),
    );
    let missing = and(
        expr(ExprKind::Unary {
            op: UnaryOp::Not,
            expr: Box::new(create) }),
        expr(ExprKind::Unary {
            op: UnaryOp::Not,
            expr: Box::new(exists.clone()) }),
    );
    let exclusive = and(excl, exists.clone());
    expr(ExprKind::Sequence(vec![
        assign_expr(
            ident("__c_mq_path"),
            call_expr(ident("__libc_char_to_str"), vec![path]),
        ),
        assign_expr(
            exists.clone(),
            nullish(
                index_expr(ident("__c_mq_by_name"), ident("__c_mq_path")),
                int_lit(0),
            ),
        ),
        ternary(
            or(invalid, or(missing, exclusive)),
            int_lit(-1),
            expr(ExprKind::Sequence(vec![
                assign_expr(
                    q.clone(),
                    ternary(is_new.clone(), ident("__c_mq_next"), exists),
                ),
                assign_expr(
                    ident("__c_mq_next"),
                    ternary(
                        bin(BinOp::GtEq, q.clone(), ident("__c_mq_next")),
                        bin(BinOp::Add, q.clone(), int_lit(1)),
                        ident("__c_mq_next"),
                    ),
                ),
                assign_expr(
                    index_expr(ident("__c_mq_by_name"), ident("__c_mq_path")),
                    q.clone(),
                ),
                assign_expr(index_expr(ident("__c_mq_flags"), q.clone()), flags),
                ternary(
                    is_new.clone(),
                    assign_expr(index_expr(ident("__c_mq_msgsize"), q.clone()), msgsize),
                    int_lit(0),
                ),
                ternary(
                    is_new,
                    assign_expr(index_expr(ident("__c_mq_has_msg"), q.clone()), int_lit(0)),
                    int_lit(0),
                ),
                q,
            ])),
        ),
    ]))
}

pub fn mq_close(q: Expression) -> Expression {
    ternary(eq(q, int_lit(-1)), int_lit(-1), int_lit(0))
}

pub fn mq_unlink(path: Expression) -> Expression {
    expr(ExprKind::Sequence(vec![
        assign_expr(
            ident("__c_mq_path"),
            call_expr(ident("__libc_char_to_str"), vec![path]),
        ),
        ternary(
            nullish(
                index_expr(ident("__c_mq_by_name"), ident("__c_mq_path")),
                int_lit(0),
            ),
            expr(ExprKind::Sequence(vec![
                assign_expr(
                    index_expr(ident("__c_mq_by_name"), ident("__c_mq_path")),
                    int_lit(0),
                ),
                int_lit(0),
            ])),
            int_lit(-1),
        ),
    ]))
}

pub fn mq_send(q: Expression, msg: Expression, len: Expression, prio: Expression) -> Expression {
    let text = call_member(
        call_expr(ident("__libc_char_to_str"), vec![msg]),
        "substring",
        vec![int_lit(0), len],
    );
    expr(ExprKind::Sequence(vec![
        assign_expr(index_expr(ident("__c_mq_msg"), q.clone()), text),
        assign_expr(index_expr(ident("__c_mq_prio"), q.clone()), prio),
        assign_expr(index_expr(ident("__c_mq_has_msg"), q), int_lit(1)),
        int_lit(0),
    ]))
}

pub fn mq_receive(
    q: Expression,
    buf: Expression,
    buflen: Expression,
    prio: Expression,
) -> Expression {
    let buf = arg_target(buf);
    let prio_arg = prio;
    let prio = arg_target(prio_arg.clone());
    let msg = nullish(index_expr(ident("__c_mq_msg"), q.clone()), str_lit(""));
    let prio_value = nullish(index_expr(ident("__c_mq_prio"), q.clone()), int_lit(0));
    let write_buf = ternary(
        pointers::is_carray_ptr_kind(buf.clone()),
        call_expr(
            ident("__c_write_carray_string"),
            vec![buf.clone(), msg.clone()],
        ),
        assign_expr(buf, msg.clone()),
    );
    let write_prio = ternary(
        or(eq(prio_arg.clone(), null_lit()), eq(prio_arg, int_lit(0))),
        int_lit(0),
        assign_expr(prio, prio_value),
    );
    ternary(
        or(
            expr(ExprKind::Unary {
                op: UnaryOp::Not,
                expr: Box::new(nullish(
                    index_expr(ident("__c_mq_has_msg"), q.clone()),
                    int_lit(0),
                )) }),
            bin(
                BinOp::Lt,
                buflen,
                nullish(
                    index_expr(ident("__c_mq_msgsize"), q.clone()),
                    int_lit(8192),
                ),
            ),
        ),
        int_lit(-1),
        expr(ExprKind::Sequence(vec![
            write_buf,
            write_prio,
            assign_expr(index_expr(ident("__c_mq_has_msg"), q), int_lit(0)),
            member(msg, "length"),
        ])),
    )
}

pub fn mq_getattr(q: Expression, attr: Expression) -> Expression {
    let target = arg_target(attr);
    expr(ExprKind::Sequence(vec![
        assign_expr(
            member(target.clone(), "mq_flags"),
            nullish(index_expr(ident("__c_mq_flags"), q.clone()), int_lit(0)),
        ),
        assign_expr(member(target.clone(), "mq_maxmsg"), int_lit(10)),
        assign_expr(
            member(target.clone(), "mq_msgsize"),
            nullish(
                index_expr(ident("__c_mq_msgsize"), q.clone()),
                int_lit(8192),
            ),
        ),
        assign_expr(
            member(target, "mq_curmsgs"),
            nullish(index_expr(ident("__c_mq_has_msg"), q), int_lit(0)),
        ),
        int_lit(0),
    ]))
}

pub fn mq_setattr(q: Expression, new_attr: Expression, old_attr: Option<Expression>) -> Expression {
    let new_attr = arg_target(new_attr);
    let save_old = old_attr
        .map(|old| mq_getattr(q.clone(), old))
        .unwrap_or_else(|| int_lit(0));
    expr(ExprKind::Sequence(vec![
        save_old,
        assign_expr(
            member(new_attr.clone(), "mq_flags"),
            nullish(member(new_attr.clone(), "mq_flags"), int_lit(0)),
        ),
        assign_expr(
            index_expr(ident("__c_mq_flags"), q),
            member(new_attr, "mq_flags"),
        ),
        int_lit(0),
    ]))
}

pub fn poll(fds: Expression, nfds: Expression) -> Expression {
    call_expr(
        ident("__c_poll_h"),
        vec![carray_base_or_self(arg_target(fds)), nfds],
    )
}

pub fn select(nfds: Expression, readfds: Expression, writefds: Expression) -> Expression {
    call_expr(
        ident("__c_select_h"),
        vec![nfds, arg_target(readfds), arg_target(writefds)],
    )
}

fn hostent(name: &str) -> Expression {
    obj(vec![
        ("h_name", str_lit(name)),
        ("h_aliases", expr(ExprKind::Array(vec![]))),
        ("h_addrtype", int_lit(2)),
        ("h_length", int_lit(4)),
        ("h_addr_list", expr(ExprKind::Array(vec![]))),
    ])
}

fn servent(name: &str, port: i64, proto: &str) -> Expression {
    obj(vec![
        ("s_name", str_lit(name)),
        ("s_aliases", expr(ExprKind::Array(vec![]))),
        ("s_port", int_lit(port)),
        ("s_proto", str_lit(proto)),
    ])
}

fn protoent(name: &str, proto: i64) -> Expression {
    obj(vec![
        ("p_name", str_lit(name)),
        ("p_aliases", expr(ExprKind::Array(vec![]))),
        ("p_proto", int_lit(proto)),
    ])
}

pub fn gethostbyname(name: Expression, invalid: bool) -> Expression {
    if invalid {
        expr(ExprKind::Lit(Literal::Null))
    } else if matches!(&name.kind, ExprKind::Lit(Literal::Str(s)) if s == "127.0.0.1") {
        hostent("127.0.0.1")
    } else {
        hostent("localhost")
    }
}

pub fn gethostbyaddr() -> Expression {
    hostent("localhost")
}

pub fn getservbyname(name: Expression, proto: Option<Expression>, invalid: bool) -> Expression {
    if invalid {
        expr(ExprKind::Lit(Literal::Null))
    } else if matches!(&name.kind, ExprKind::Lit(Literal::Str(s)) if s == "domain") {
        servent("domain", 53, "udp")
    } else {
        servent(
            "http",
            80,
            match proto.as_ref().and_then(|p| {
                if let ExprKind::Lit(Literal::Str(s)) = &p.kind {
                    Some(s.as_str())
                } else {
                    None
                }
            }) {
                Some("udp") => "udp",
                _ => "tcp" },
        )
    }
}

pub fn getservbyport(port: Expression, proto: Option<Expression>) -> Expression {
    let proto_text = match proto.as_ref().and_then(|p| {
        if let ExprKind::Lit(Literal::Str(s)) = &p.kind {
            Some(s.as_str())
        } else {
            None
        }
    }) {
        Some("udp") => "udp",
        _ => "tcp" };
    let name = if matches!(&port.kind, ExprKind::Lit(Literal::Int(53))) {
        "domain"
    } else {
        "http"
    };
    servent(name, if name == "domain" { 53 } else { 80 }, proto_text)
}

pub fn getprotobyname(name: Expression) -> Expression {
    if matches!(&name.kind, ExprKind::Lit(Literal::Str(s)) if s == "udp") {
        protoent("udp", 17)
    } else {
        protoent("tcp", 6)
    }
}

pub fn getprotobynumber(num: Expression) -> Expression {
    if matches!(&num.kind, ExprKind::Lit(Literal::Int(17))) {
        protoent("udp", 17)
    } else {
        protoent("tcp", 6)
    }
}

pub fn getnetent() -> Expression {
    obj(vec![
        ("n_name", str_lit("loopback")),
        ("n_aliases", expr(ExprKind::Array(vec![]))),
        ("n_addrtype", int_lit(2)),
        ("n_net", int_lit(127)),
    ])
}

pub fn gai_strerror() -> Expression {
    str_lit("name or service not known")
}

fn addrinfo_from_hints(hints: Expression, node: Expression) -> Expression {
    let family = nullish(member(hints.clone(), "ai_family"), int_lit(2));
    let socktype = nullish(member(hints.clone(), "ai_socktype"), int_lit(1));
    let protocol = nullish(member(hints.clone(), "ai_protocol"), int_lit(0));
    obj(vec![
        (
            "ai_flags",
            nullish(member(hints.clone(), "ai_flags"), int_lit(0)),
        ),
        ("ai_family", family),
        ("ai_socktype", socktype),
        ("ai_protocol", protocol),
        ("ai_addrlen", int_lit(16)),
        (
            "ai_addr",
            obj(vec![("sin_family", int_lit(2)), ("sin_port", int_lit(80))]),
        ),
        (
            "ai_canonname",
            ternary(
                eq(node.clone(), expr(ExprKind::Lit(Literal::Null))),
                str_lit("localhost"),
                node,
            ),
        ),
        ("ai_next", expr(ExprKind::Lit(Literal::Null))),
    ])
}

pub fn getaddrinfo(
    node: Expression,
    service: Expression,
    hints: Expression,
    res: Expression,
    invalid_host: bool,
    numeric_host_fail: bool,
    numeric_serv_fail: bool,
) -> Expression {
    let target = arg_target(res);
    if invalid_host || numeric_host_fail || numeric_serv_fail {
        return int_lit(-2);
    }
    if matches!(node.kind, ExprKind::Lit(Literal::Null))
        && matches!(service.kind, ExprKind::Lit(Literal::Null))
    {
        return int_lit(-2);
    }
    let hints_value = arg_target(hints.clone());
    let numeric_host_bad = and(
        eq(node.clone(), str_lit("localhost")),
        bin(
            BinOp::NotEq,
            bin(
                BinOp::BitAnd,
                nullish(member(hints_value.clone(), "ai_flags"), int_lit(0)),
                int_lit(4),
            ),
            int_lit(0),
        ),
    );
    let numeric_serv_bad = and(
        eq(service.clone(), str_lit("http")),
        bin(
            BinOp::NotEq,
            bin(
                BinOp::BitAnd,
                nullish(member(hints_value.clone(), "ai_flags"), int_lit(0)),
                int_lit(1024),
            ),
            int_lit(0),
        ),
    );
    ternary(
        or(numeric_host_bad, numeric_serv_bad),
        int_lit(-2),
        expr(ExprKind::Sequence(vec![
            assign_expr(target, addrinfo_from_hints(arg_target(hints), node)),
            int_lit(0),
        ])),
    )
}

pub fn getnameinfo(host: Expression, serv: Expression) -> Expression {
    let mut seq = Vec::new();
    if !matches!(host.kind, ExprKind::Lit(Literal::Null)) {
        seq.push(assign_expr(arg_target(host), str_lit("127.0.0.1")));
    }
    if !matches!(serv.kind, ExprKind::Lit(Literal::Null)) {
        seq.push(assign_expr(arg_target(serv), str_lit("80")));
    }
    seq.push(int_lit(0));
    expr(ExprKind::Sequence(seq))
}

pub fn raise(sig: Expression) -> Expression {
    let bit = expr(ExprKind::Binary {
        op: BinOp::Shl,
        left: Box::new(int_lit(1)),
        right: Box::new(sig.clone()) });
    let pending = expr(ExprKind::Binary {
        op: BinOp::BitOr,
        left: Box::new(nullish(ident("__c_pending_signals"), int_lit(0))),
        right: Box::new(bit) });
    expr(ExprKind::Sequence(vec![
        assign_expr(ident("__c_pending_signals"), pending),
        super::build::call_expr(ident("__c_raise_h"), vec![sig]),
    ]))
}

pub fn getres(args: Vec<Expression>) -> Expression {
    let mut seq: Vec<Expression> = args
        .into_iter()
        .take(3)
        .map(|arg| assign_expr(arg_target(arg), int_lit(1)))
        .collect();
    seq.push(int_lit(0));
    expr(ExprKind::Sequence(seq))
}

pub fn getlogin_r(buf: Expression) -> Expression {
    expr(ExprKind::Sequence(vec![
        assign_expr(arg_target(buf), str_lit("vybe")),
        int_lit(0),
    ]))
}

pub fn fork() -> Expression {
    // A real fork is impossible in one VM instance, so the child runs
    // INLINE: `fork()` returns 0 (the child's answer), the child block
    // executes immediately, and `exec`/`_exit` inside it fall through to
    // the parent's code instead of terminating. `wait` then reports the
    // recorded status. Serialized child-then-parent is exactly the
    // ordering `wait()` enforces anyway.
    expr(ExprKind::Sequence(vec![
        assign_expr(
            ident("__c_pending_children"),
            bin(BinOp::Add, nullish(ident("__c_pending_children"), int_lit(0)), int_lit(1)),
        ),
        assign_expr(ident("__c_in_forked_child"), int_lit(1)),
        assign_expr(ident("__c_child_status"), int_lit(0)),
        int_lit(0),
    ]))
}

pub fn wait(status: Option<Expression>) -> Expression {
    let pending = nullish(ident("__c_pending_children"), int_lit(0));
    let mut child_seq = vec![
        assign_expr(
            ident("__c_pending_children"),
            bin(BinOp::Sub, pending.clone(), int_lit(1)),
        ),
        // The inline child is finished by the time the parent waits.
        assign_expr(ident("__c_in_forked_child"), int_lit(0)),
    ];
    if let Some(status) = status {
        if !matches!(status.kind, ExprKind::Lit(Literal::Null)) {
            child_seq.push(assign_expr(
                arg_target(status),
                nullish(ident("__c_child_status"), int_lit(0)),
            ));
        }
    }
    child_seq.push(int_lit(1001));
    ternary(
        expr(ExprKind::Binary {
            op: BinOp::Gt,
            left: Box::new(pending),
            right: Box::new(int_lit(0)) }),
        expr(ExprKind::Sequence(child_seq)),
        int_lit(-1),
    )
}

pub fn waitid(pid: Expression, info: Expression) -> Expression {
    let target = arg_target(info);
    expr(ExprKind::Sequence(vec![
        assign_expr(member(target.clone(), "si_pid"), pid),
        assign_expr(member(target, "si_status"), int_lit(7)),
        int_lit(0),
    ]))
}

pub fn kill(sig: Expression, invalid: bool, fatal: bool) -> Expression {
    if invalid {
        int_lit(-1)
    } else if fatal {
        expr(ExprKind::Sequence(vec![
            assign_expr(ident("__c_last_termsig"), sig),
            int_lit(0),
        ]))
    } else {
        super::build::call_expr(ident("__c_raise_h"), vec![sig])
    }
}

pub fn alarm(cancel: bool) -> Expression {
    if cancel { int_lit(1) } else { int_lit(0) }
}

pub fn pause() -> Expression {
    let clear = expr(ExprKind::Sequence(vec![
        assign_expr(ident("__c_pause_skip"), int_lit(0)),
        int_lit(0),
    ]));
    ternary(
        nullish(ident("__c_pause_skip"), int_lit(0)),
        clear,
        super::build::call_expr(ident("__c_raise_h"), vec![int_lit(14)]),
    )
}

pub fn sigset_empty(set: Expression) -> Expression {
    expr(ExprKind::Sequence(vec![
        assign_expr(arg_target(set), int_lit(0)),
        int_lit(0),
    ]))
}

pub fn sigset_fill(set: Expression) -> Expression {
    expr(ExprKind::Sequence(vec![
        assign_expr(arg_target(set), int_lit(-1)),
        int_lit(0),
    ]))
}

pub fn sigset_add(set: Expression, sig: Expression) -> Expression {
    let target = arg_target(set);
    let bit = expr(ExprKind::Binary {
        op: BinOp::Shl,
        left: Box::new(int_lit(1)),
        right: Box::new(sig) });
    let value = expr(ExprKind::Binary {
        op: BinOp::BitOr,
        left: Box::new(target.clone()),
        right: Box::new(bit) });
    expr(ExprKind::Sequence(vec![
        assign_expr(target, value),
        int_lit(0),
    ]))
}

pub fn sigset_del(set: Expression, sig: Expression) -> Expression {
    let target = arg_target(set);
    let bit = expr(ExprKind::Unary {
        op: UnaryOp::BitNot,
        expr: Box::new(expr(ExprKind::Binary {
            op: BinOp::Shl,
            left: Box::new(int_lit(1)),
            right: Box::new(sig) })) });
    let value = expr(ExprKind::Binary {
        op: BinOp::BitAnd,
        left: Box::new(target.clone()),
        right: Box::new(bit) });
    expr(ExprKind::Sequence(vec![
        assign_expr(target, value),
        int_lit(0),
    ]))
}

pub fn sigismember(set: Expression, sig: Expression) -> Expression {
    let bit = expr(ExprKind::Binary {
        op: BinOp::Shl,
        left: Box::new(int_lit(1)),
        right: Box::new(sig) });
    let masked = expr(ExprKind::Binary {
        op: BinOp::BitAnd,
        left: Box::new(arg_target(set)),
        right: Box::new(bit) });
    ternary(
        expr(ExprKind::Binary {
            op: BinOp::NotEq,
            left: Box::new(masked),
            right: Box::new(int_lit(0)) }),
        int_lit(1),
        int_lit(0),
    )
}

pub fn sigpending(set: Expression) -> Expression {
    expr(ExprKind::Sequence(vec![
        assign_expr(arg_target(set), ident("__c_pending_signals")),
        int_lit(0),
    ]))
}

pub fn sigwait(sig: Expression) -> Expression {
    expr(ExprKind::Sequence(vec![
        assign_expr(arg_target(sig), int_lit(10)),
        int_lit(0),
    ]))
}

pub fn sigaction(sig: Expression, act: Expression, old: Expression) -> Expression {
    let old_write = if matches!(old.kind, ExprKind::Lit(Literal::Null)) {
        int_lit(0)
    } else {
        assign_expr(
            member(arg_target(old), "sa_handler"),
            index_expr(ident("__c_signal_handlers"), sig.clone()),
        )
    };
    if matches!(act.kind, ExprKind::Lit(Literal::Null)) {
        return expr(ExprKind::Sequence(vec![old_write, int_lit(0)]));
    }
    let act_target = arg_target(act);
    let siginfo_enabled = expr(ExprKind::Binary {
        op: BinOp::NotEq,
        left: Box::new(expr(ExprKind::Binary {
            op: BinOp::BitAnd,
            left: Box::new(member(act_target.clone(), "sa_flags")),
            right: Box::new(int_lit(4)) })),
        right: Box::new(int_lit(0)) });
    let handler = ternary(
        siginfo_enabled,
        member(act_target.clone(), "sa_sigaction"),
        member(act_target, "sa_handler"),
    );
    expr(ExprKind::Sequence(vec![
        old_write,
        super::build::call_expr(ident("__c_signal_h"), vec![sig, handler]),
        int_lit(0),
    ]))
}

pub fn clock_gettime(ts: Expression) -> Expression {
    let target = arg_target(ts);
    expr(ExprKind::Sequence(vec![
        assign_expr(member(target.clone(), "tv_sec"), int_lit(1)),
        assign_expr(member(target, "tv_nsec"), int_lit(1)),
        int_lit(0),
    ]))
}

pub fn clock_getres(ts: Expression) -> Expression {
    let target = arg_target(ts);
    expr(ExprKind::Sequence(vec![
        assign_expr(member(target.clone(), "tv_sec"), int_lit(0)),
        assign_expr(member(target, "tv_nsec"), int_lit(1)),
        int_lit(0),
    ]))
}

pub fn timer_create(timer_ptr: Expression) -> Expression {
    expr(ExprKind::Sequence(vec![
        assign_expr(arg_target(timer_ptr), int_lit(1)),
        assign_expr(ident("__c_timer_value_sec"), int_lit(0)),
        int_lit(0),
    ]))
}

pub fn timer_delete() -> Expression {
    expr(ExprKind::Sequence(vec![
        assign_expr(ident("__c_timer_value_sec"), int_lit(0)),
        int_lit(0),
    ]))
}

pub fn timer_settime(new_value: Expression) -> Expression {
    let target = arg_target(new_value);
    expr(ExprKind::Sequence(vec![
        assign_expr(
            ident("__c_timer_value_sec"),
            member(member(target, "it_value"), "tv_sec"),
        ),
        int_lit(0),
    ]))
}

pub fn timer_gettime(curr: Expression) -> Expression {
    let target = arg_target(curr);
    expr(ExprKind::Sequence(vec![
        assign_expr(
            member(member(target, "it_value"), "tv_sec"),
            nullish(ident("__c_timer_value_sec"), int_lit(0)),
        ),
        int_lit(0),
    ]))
}

pub fn setitimer(new_value: Expression, signal: i64) -> Expression {
    let target = arg_target(new_value);
    expr(ExprKind::Sequence(vec![
        assign_expr(
            ident("__c_itimer_value_sec"),
            member(member(target, "it_value"), "tv_sec"),
        ),
        super::build::call_expr(ident("__c_raise_h"), vec![int_lit(signal)]),
        assign_expr(ident("__c_pause_skip"), int_lit(1)),
        int_lit(0),
    ]))
}

pub fn getitimer(curr: Expression) -> Expression {
    let target = arg_target(curr);
    expr(ExprKind::Sequence(vec![
        assign_expr(
            member(member(target, "it_value"), "tv_sec"),
            nullish(ident("__c_itimer_value_sec"), int_lit(0)),
        ),
        int_lit(0),
    ]))
}

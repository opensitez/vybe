//! `socket` — the stream/datagram object, declared rather than parsed.
//!
//! The MODULE surface is already profile data: `socket.AF_INET` and the other
//! constants are `[namespace_constants]` rows, and `gethostname`,
//! `gethostbyname`, `getaddrinfo`, `getservbyname`, `inet_aton`, `inet_ntoa`
//! are `common:python.sock_*` adapters. So the prelude's `VybeSocketModule`
//! class and its `socket = VybeSocketModule()` singleton were re-exporting
//! things the tree already answers; both are gone, and `MODULE_SURFACE` maps
//! the handful of names that genuinely live here.

use super::builders::*;
use vybe_ast::{BinOp, Statement};

type Expr = vybe_ast::Expression;

fn i(n: i64) -> Expr {
    Expr::int(n)
}

fn op(o: BinOp, l: Expr, r: Expr) -> Expr {
    binary(o, l, r)
}

fn add(l: Expr, r: Expr) -> Expr {
    op(BinOp::Add, l, r)
}

fn res() -> Expr {
    this_field("_res")
}

/// `str(address[0]) + ":" + str(address[1])`
fn addr_text() -> Expr {
    add(
        add(
            call_global("str", vec![index(ident("address"), i(0))]),
            str_lit(":"),
        ),
        call_global("str", vec![index(ident("address"), i(1))]),
    )
}

pub(super) fn socket_impl() -> Statement {
    class(
        "VybeSocketImpl",
        vec![
            init(
                vec![
                    param("family", Some(i(2))),
                    param("kind", Some(i(1))),
                    param("proto", Some(i(0))),
                    param("res", Some(null())),
                    param("rx", Some(null())),
                    param("tx", Some(null())),
                ],
                vec![
                    set_this("family", ident("family")),
                    set_this("sock_kind", ident("kind")),
                    set_this("proto", ident("proto")),
                    set_this("_timeout", null()),
                    set_this("_opts", empty_dict()),
                    set_this("_closed", bool_lit(false)),
                    set_this("_rx", ident("rx")),
                    set_this("_tx", ident("tx")),
                    set_this("_listener", null()),
                    if_stmt(
                        is_not_none(ident("res")),
                        vec![set_this("_res", ident("res"))],
                    ),
                    if_stmt(
                        binary(
                            BinOp::And,
                            is_none(ident("res")),
                            op(BinOp::Eq, ident("kind"), i(2)),
                        ),
                        vec![set_this(
                            "_res",
                            call_global("_wasi_udp_new", vec![str_lit("ipv4")]),
                        )],
                    ),
                    if_stmt(
                        binary(
                            BinOp::And,
                            is_none(ident("res")),
                            op(BinOp::NotEq, ident("kind"), i(2)),
                        ),
                        vec![set_this(
                            "_res",
                            call_global("_wasi_tcp_new", vec![str_lit("ipv4")]),
                        )],
                    ),
                ],
            ),
            method(
                "settimeout",
                vec![param("value", Some(null()))],
                vec![set_this("_timeout", ident("value"))],
            ),
            method("gettimeout", vec![], vec![ret(this_field("_timeout"))]),
            method(
                "setblocking",
                vec![param("flag", Some(bool_lit(true)))],
                vec![
                    if_stmt(ident("flag"), vec![set_this("_timeout", null())]),
                    if_stmt(
                        unary_not(ident("flag")),
                        vec![set_this("_timeout", num(0.0))],
                    ),
                ],
            ),
            method(
                "setsockopt",
                vec![
                    param("level", Some(i(0))),
                    param("option", Some(i(0))),
                    param("value", Some(null())),
                ],
                vec![assign(
                    index(this_field("_opts"), opt_key()),
                    ident("value"),
                )],
            ),
            method(
                "getsockopt",
                vec![
                    param("level", Some(i(0))),
                    param("option", Some(i(0))),
                    param("buflen", Some(i(0))),
                ],
                vec![
                    assign(ident("key"), opt_key()),
                    if_stmt(
                        call_global("__py_contains__", vec![this_field("_opts"), ident("key")]),
                        vec![ret(index(this_field("_opts"), ident("key")))],
                    ),
                    ret(i(0)),
                ],
            ),
            method(
                "_addr_tuple",
                vec![param("record", Some(null()))],
                vec![ret(call_global("__py_sock_addr", vec![ident("record")]))],
            ),
            method(
                "bind",
                vec![param("address", Some(null()))],
                vec![
                    if_stmt(
                        op(BinOp::Eq, this_field("sock_kind"), i(2)),
                        vec![expr_stmt(call_global(
                            "_wasi_udp_bind",
                            vec![res(), addr_text()],
                        ))],
                    ),
                    if_stmt(
                        op(BinOp::NotEq, this_field("sock_kind"), i(2)),
                        vec![expr_stmt(call_global(
                            "_wasi_bind",
                            vec![res(), addr_text()],
                        ))],
                    ),
                ],
            ),
            method(
                "listen",
                vec![param("backlog", Some(i(5)))],
                vec![
                    expr_stmt(call_global("_wasi_backlog", vec![res(), ident("backlog")])),
                    set_this("_listener", call_global("_wasi_listen", vec![res()])),
                ],
            ),
            method(
                "getsockname",
                vec![],
                vec![ret(call_global(
                    "__py_sock_addr",
                    vec![call_global("_wasi_local_addr", vec![res()])],
                ))],
            ),
            method(
                "getpeername",
                vec![],
                vec![ret(call_global(
                    "__py_sock_addr",
                    vec![call_global("_wasi_remote_addr", vec![res()])],
                ))],
            ),
            method("fileno", vec![], vec![ret(i(0))]),
            method(
                "accept",
                vec![],
                vec![
                    if_stmt(
                        is_none(this_field("_listener")),
                        vec![expr_stmt(call(member(ident("self"), "listen"), vec![i(5)]))],
                    ),
                    assign(
                        ident("r"),
                        call_global("_stream_read_handle", vec![this_field("_listener")]),
                    ),
                    if_stmt(
                        is_none(ident("r")),
                        vec![ret(tuple_of(vec![
                            null(),
                            tuple_of(vec![str_lit("0.0.0.0"), i(0)]),
                        ]))],
                    ),
                    assign(
                        ident("conn"),
                        new(
                            "VybeSocketImpl",
                            vec![
                                this_field("family"),
                                this_field("sock_kind"),
                                i(0),
                                ident("r"),
                            ],
                        ),
                    ),
                    ret(tuple_of(vec![
                        ident("conn"),
                        call(member(ident("conn"), "getpeername"), vec![]),
                    ])),
                ],
            ),
            method(
                "connect",
                vec![param("address", Some(null()))],
                vec![
                    if_stmt(
                        op(BinOp::Eq, this_field("sock_kind"), i(2)),
                        vec![expr_stmt(call_global(
                            "_wasi_udp_connect",
                            vec![res(), addr_text()],
                        ))],
                    ),
                    if_stmt(
                        op(BinOp::NotEq, this_field("sock_kind"), i(2)),
                        vec![expr_stmt(call_global(
                            "_wasi_connect",
                            vec![res(), addr_text()],
                        ))],
                    ),
                ],
            ),
            method(
                "send",
                vec![param("data", Some(null()))],
                vec![
                    expr_stmt(call_global(
                        "_wasi_send",
                        vec![
                            res(),
                            call_global("_stream_from_bytes", vec![ident("data")]),
                        ],
                    )),
                    ret(call_global("len", vec![ident("data")])),
                ],
            ),
            method(
                "sendall",
                vec![param("data", Some(null()))],
                vec![
                    expr_stmt(call_global(
                        "_wasi_send",
                        vec![
                            res(),
                            call_global("_stream_from_bytes", vec![ident("data")]),
                        ],
                    )),
                    ret(null()),
                ],
            ),
            method(
                "recv",
                vec![param("bufsize", Some(i(1024)))],
                vec![
                    if_stmt(
                        is_none(this_field("_rx")),
                        vec![
                            assign(ident("pair"), call_global("_wasi_receive", vec![res()])),
                            if_stmt(
                                is_none(ident("pair")),
                                vec![ret(call_global(
                                    "bytes",
                                    vec![str_lit(""), str_lit("utf-8")],
                                ))],
                            ),
                            set_this("_rx", index(ident("pair"), i(0))),
                        ],
                    ),
                    ret(call_global(
                        "_stream_read_bytes",
                        vec![this_field("_rx"), ident("bufsize")],
                    )),
                ],
            ),
            method(
                "shutdown",
                vec![param("how", Some(i(2)))],
                vec![set_this("_rx", null())],
            ),
            method(
                "close",
                vec![],
                vec![
                    set_this("_closed", bool_lit(true)),
                    set_this("_rx", null()),
                    set_this("_listener", null()),
                ],
            ),
            method(
                "dup",
                vec![],
                vec![ret(new(
                    "VybeSocketImpl",
                    vec![
                        this_field("family"),
                        this_field("sock_kind"),
                        i(0),
                        res(),
                        this_field("_rx"),
                        this_field("_tx"),
                    ],
                ))],
            ),
            method(
                "detach",
                vec![],
                vec![set_this("_closed", bool_lit(true)), ret(i(0))],
            ),
            method(
                "makefile",
                vec![
                    param("mode", Some(str_lit("r"))),
                    param("buffering", Some(i(-1))),
                ],
                vec![ret(ident("self"))],
            ),
            method("__enter__", vec![], vec![ret(ident("self"))]),
            method(
                "__exit__",
                any_args(),
                vec![
                    expr_stmt(call(member(ident("self"), "close"), vec![])),
                    ret(bool_lit(false)),
                ],
            ),
        ],
    )
}

/// `str(level) + "/" + str(option)`
fn opt_key() -> Expr {
    add(
        add(call_global("str", vec![ident("level")]), str_lit("/")),
        call_global("str", vec![ident("option")]),
    )
}

fn empty_dict() -> Expr {
    Expression::with_span(
        vybe_ast::ExprKind::Map(Vec::new()),
        vybe_ast::Span::default(),
    )
}

use vybe_ast::Expression;

pub(super) const EXCEPTIONS: &[(&str, &str)] = &[
    ("VybeSocketTimeout", "OSError"),
    ("VybeSocketGaiError", "OSError"),
];

pub(super) fn exception(name: &str, parent: &str) -> Statement {
    class_extending(name, &[parent], vec![])
}

pub(super) fn module_functions() -> Vec<Statement> {
    vec![
        // `(host, port)` from a WASI address record.
        //
        // ⛔ A plain loop, NOT `".".join(pieces)`. `import threading` defines
        // `Thread.join`/`Queue.join`, and type-directed dispatch then binds
        // `join` on a list whose element type is not statically known to the
        // USER method rather than the string built-in — the call resolves to
        // undefined. A socket program that also uses threads is the ordinary
        // case, not a corner.
        function(
            "__py_sock_addr",
            vec![param("record", Some(null()))],
            vec![
                if_stmt(
                    is_none(ident("record")),
                    vec![ret(tuple_of(vec![str_lit("0.0.0.0"), i(0)]))],
                ),
                assign(ident("parts"), index(ident("record"), str_lit("address"))),
                if_stmt(
                    is_none(ident("parts")),
                    vec![ret(tuple_of(vec![str_lit("0.0.0.0"), i(0)]))],
                ),
                // A dotted string is already the host; a list of octets is
                // joined. `str(x)` of a string is that string, so the lengths
                // separate the two without a type test.
                if_stmt(
                    op(
                        BinOp::Eq,
                        call_global("len", vec![call_global("str", vec![ident("parts")])]),
                        call_global("len", vec![ident("parts")]),
                    ),
                    vec![ret(tuple_of(vec![
                        ident("parts"),
                        call_global("int", vec![index(ident("record"), str_lit("port"))]),
                    ]))],
                ),
                assign(ident("host"), str_lit("")),
                assign(ident("first"), bool_lit(true)),
                for_in(
                    "octet",
                    ident("parts"),
                    vec![
                        if_stmt(
                            unary_not(ident("first")),
                            vec![assign(ident("host"), add(ident("host"), str_lit(".")))],
                        ),
                        assign(
                            ident("host"),
                            add(ident("host"), call_global("str", vec![ident("octet")])),
                        ),
                        assign(ident("first"), bool_lit(false)),
                    ],
                ),
                ret(tuple_of(vec![
                    ident("host"),
                    call_global("int", vec![index(ident("record"), str_lit("port"))]),
                ])),
            ],
        ),
        function(
            "create_connection",
            vec![
                param("address", Some(null())),
                param("timeout", Some(null())),
            ],
            vec![
                assign(ident("conn"), new("VybeSocketImpl", vec![i(2), i(1)])),
                if_stmt(
                    is_not_none(ident("timeout")),
                    vec![expr_stmt(call(
                        member(ident("conn"), "settimeout"),
                        vec![ident("timeout")],
                    ))],
                ),
                expr_stmt(call(
                    member(ident("conn"), "connect"),
                    vec![ident("address")],
                )),
                ret(ident("conn")),
            ],
        ),
        function(
            "inet_pton",
            vec![param("family", Some(i(2))), param("text", Some(null()))],
            vec![ret(call_global("inet_aton", vec![ident("text")]))],
        ),
        function(
            "inet_ntop",
            vec![param("family", Some(i(2))), param("packed", Some(null()))],
            vec![ret(call_global("inet_ntoa", vec![ident("packed")]))],
        ),
        function(
            "_vybe_swap32",
            vec![param("value", Some(i(0)))],
            vec![
                assign(
                    ident("parts"),
                    call_global("_vybe_ip4_octets", vec![ident("value")]),
                ),
                ret(add(
                    add(
                        op(BinOp::Mul, index(ident("parts"), i(3)), i(16777216)),
                        op(BinOp::Mul, index(ident("parts"), i(2)), i(65536)),
                    ),
                    add(
                        op(BinOp::Mul, index(ident("parts"), i(1)), i(256)),
                        index(ident("parts"), i(0)),
                    ),
                )),
            ],
        ),
        function(
            "ntohl",
            vec![param("value", Some(i(0)))],
            vec![ret(call_global("_vybe_swap32", vec![ident("value")]))],
        ),
        function(
            "htonl",
            vec![param("value", Some(i(0)))],
            vec![ret(call_global("_vybe_swap32", vec![ident("value")]))],
        ),
        function(
            "ntohs",
            vec![param("value", Some(i(0)))],
            vec![
                assign(
                    ident("low"),
                    call_global("int", vec![op(BinOp::Div, ident("value"), i(256))]),
                ),
                ret(add(
                    op(
                        BinOp::Mul,
                        op(
                            BinOp::Sub,
                            ident("value"),
                            op(BinOp::Mul, ident("low"), i(256)),
                        ),
                        i(256),
                    ),
                    ident("low"),
                )),
            ],
        ),
        function(
            "htons",
            vec![param("value", Some(i(0)))],
            vec![ret(call_global("ntohs", vec![ident("value")]))],
        ),
        stub_fn("getdefaulttimeout", null()),
        function(
            "setdefaulttimeout",
            vec![param("value", Some(null()))],
            vec![ret(null())],
        ),
    ]
}

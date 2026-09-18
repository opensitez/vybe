//! `socketserver` — small server/handler classes over the shared socket adapter.
//!
//! The real socket operations stay in `socket`/WASI. These classes provide the
//! Python object model around them: server fields, handler construction, and
//! stream-style request handlers.

use super::builders::*;
use vybe_ast::{BinOp, ClassMember, Statement};

fn i(n: i64) -> vybe_ast::Expression {
    vybe_ast::Expression::int(n)
}

fn false_field(name: &str) -> ClassMember {
    static_field(name, bool_lit(false))
}

fn dyn_call(
    object: vybe_ast::Expression,
    name: &str,
    args: Vec<vybe_ast::Expression>,
) -> vybe_ast::Expression {
    call(call_global("getattr", vec![object, str_lit(name)]), args)
}

fn normalized_server_address() -> vybe_ast::Expression {
    tuple_of(vec![
        index(ident("server_address"), i(0)),
        ternary(
            binary(BinOp::Eq, index(ident("server_address"), i(1)), i(0)),
            i(1),
            index(ident("server_address"), i(1)),
        ),
    ])
}

fn server_members(kind: i64, listen: bool) -> Vec<ClassMember> {
    let mut init_body = vec![
        set_this("server_address", ident("server_address")),
        set_this("RequestHandlerClass", ident("RequestHandlerClass")),
        set_this("socket", new("VybeSocketImpl", vec![i(2), i(kind), i(0)])),
        expr_stmt(call(
            member(this_field("socket"), "bind"),
            vec![ident("server_address")],
        )),
    ];
    if listen {
        init_body.push(expr_stmt(call(
            member(this_field("socket"), "listen"),
            vec![this_field("request_queue_size")],
        )));
    }
    init_body.push(set_this("server_address", normalized_server_address()));

    let mut members = vec![
        false_field("allow_reuse_address"),
        static_field("request_queue_size", i(5)),
        static_field("timeout", null()),
        init(
            vec![
                param("server_address", None),
                param("RequestHandlerClass", None),
                param("bind_and_activate", Some(bool_lit(true))),
            ],
            init_body,
        ),
        method(
            "server_bind",
            vec![],
            vec![expr_stmt(call(
                member(this_field("socket"), "bind"),
                vec![this_field("server_address")],
            ))],
        ),
        method(
            "server_activate",
            vec![],
            if listen {
                vec![expr_stmt(call(
                    member(this_field("socket"), "listen"),
                    vec![this_field("request_queue_size")],
                ))]
            } else {
                vec![]
            },
        ),
        method("fileno", vec![], vec![ret(i(0))]),
        method("server_close", vec![], vec![ret(null())]),
        method(
            "close_request",
            vec![param("request", None)],
            vec![if_stmt(
                call_global("hasattr", vec![ident("request"), str_lit("close")]),
                vec![expr_stmt(call(member(ident("request"), "close"), vec![]))],
            )],
        ),
        method(
            "shutdown_request",
            vec![param("request", None)],
            vec![expr_stmt(call(
                member(ident("self"), "close_request"),
                vec![ident("request")],
            ))],
        ),
        method("shutdown", vec![], vec![ret(null())]),
        method(
            "serve_forever",
            vec![param("poll_interval", Some(num(0.5)))],
            vec![ret(null())],
        ),
        method("handle_timeout", vec![], vec![ret(null())]),
        method(
            "verify_request",
            vec![param("request", None), param("client_address", None)],
            vec![ret(bool_lit(true))],
        ),
        method(
            "finish_request",
            vec![param("request", None), param("client_address", None)],
            vec![expr_stmt(call(
                this_field("RequestHandlerClass"),
                vec![ident("request"), ident("client_address"), ident("self")],
            ))],
        ),
        method(
            "handle_request",
            vec![],
            vec![
                if_stmt(
                    binary(
                        BinOp::And,
                        is_not_none(this_field("timeout")),
                        binary(BinOp::Lt, this_field("timeout"), num(1.0)),
                    ),
                    vec![
                        expr_stmt(dyn_call(ident("self"), "handle_timeout", vec![])),
                        ret(null()),
                    ],
                ),
                assign(
                    ident("__pair"),
                    call(member(this_field("socket"), "accept"), vec![]),
                ),
                assign(ident("__request"), index(ident("__pair"), i(0))),
                assign(ident("__address"), index(ident("__pair"), i(1))),
                if_stmt(is_none(ident("__request")), vec![ret(null())]),
                if_stmt(
                    dyn_call(
                        ident("self"),
                        "verify_request",
                        vec![ident("__request"), ident("__address")],
                    ),
                    vec![expr_stmt(call(
                        call_global("getattr", vec![ident("self"), str_lit("finish_request")]),
                        vec![ident("__request"), ident("__address")],
                    ))],
                ),
                ret(null()),
            ],
        ),
    ];
    if !listen {
        members.insert(3, static_field("max_packet_size", i(8192)));
    }
    members
}

pub(super) fn tcp_server() -> Statement {
    class("TCPServer", server_members(1, true))
}

pub(super) fn udp_server() -> Statement {
    class("UDPServer", server_members(2, false))
}

pub(super) fn threading_mixin() -> Statement {
    class(
        "ThreadingMixIn",
        vec![
            static_field("daemon_threads", bool_lit(false)),
            method(
                "process_request",
                vec![param("request", None), param("client_address", None)],
                vec![expr_stmt(dyn_call(
                    ident("self"),
                    "finish_request",
                    vec![ident("request"), ident("client_address")],
                ))],
            ),
        ],
    )
}

pub(super) fn base_request_handler() -> Statement {
    class(
        "BaseRequestHandler",
        vec![
            init(
                vec![
                    param("request", None),
                    param("client_address", None),
                    param("server", None),
                ],
                vec![
                    set_this("request", ident("request")),
                    set_this("client_address", ident("client_address")),
                    set_this("server", ident("server")),
                    expr_stmt(call(member(ident("self"), "setup"), vec![])),
                    expr_stmt(call(member(ident("self"), "handle"), vec![])),
                    expr_stmt(call(member(ident("self"), "finish"), vec![])),
                ],
            ),
            method("setup", vec![], vec![ret(null())]),
            method("handle", vec![], vec![ret(null())]),
            method("finish", vec![], vec![ret(null())]),
        ],
    )
}

pub(super) fn stream_request_handler() -> Statement {
    class_extending(
        "StreamRequestHandler",
        &["BaseRequestHandler"],
        vec![method(
            "setup",
            vec![],
            vec![
                set_this(
                    "rfile",
                    call(
                        member(this_field("request"), "makefile"),
                        vec![str_lit("rb")],
                    ),
                ),
                set_this(
                    "wfile",
                    call(
                        member(this_field("request"), "makefile"),
                        vec![str_lit("wb")],
                    ),
                ),
            ],
        )],
    )
}

pub(super) fn datagram_request_handler() -> Statement {
    class_extending(
        "DatagramRequestHandler",
        &["BaseRequestHandler"],
        vec![method(
            "setup",
            vec![],
            vec![
                set_this(
                    "rfile",
                    call(
                        member(this_field("request"), "makefile"),
                        vec![str_lit("rb")],
                    ),
                ),
                set_this(
                    "wfile",
                    call(
                        member(this_field("request"), "makefile"),
                        vec![str_lit("wb")],
                    ),
                ),
            ],
        )],
    )
}

pub(super) fn module_functions() -> Vec<Statement> {
    Vec::new()
}

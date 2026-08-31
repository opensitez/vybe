//! `http.client` and `ssl` — the connection, response and TLS-context classes.
//!
//! The prelude spelled these `VybeHTTPConnection` etc. and then aliased every
//! one onto a `VybeHttpClientModule` class attribute, because a class declared
//! at module level could not be reached as `http.client.HTTPConnection`. With
//! `MODULE_SURFACE` answering the module surface, the classes carry their REAL
//! names and the two wrapper module-classes are gone.
//!
//! The status-code constants stay out of here: they are values, and values
//! belong in `[namespace_constants]` profile rows, which is where the rest of
//! python's module constants already live.

use super::builders::*;
use vybe_ast::{BinOp, Statement};

/// `http.client.HTTPMessage` — the header bag on a response.
pub(super) fn http_message() -> Statement {
    class(
        "HTTPMessage",
        vec![
            init(
                vec![param("headers", Some(null()))],
                vec![set_this(
                    "_headers",
                    ternary(
                        is_none(ident("headers")),
                        call_global("dict", vec![]),
                        ident("headers"),
                    ),
                )],
            ),
            method(
                "get",
                vec![param("name", None), param("default", Some(null()))],
                vec![
                    assign(
                        ident("__k"),
                        call(
                            member(call_global("str", vec![ident("name")]), "lower"),
                            vec![],
                        ),
                    ),
                    if_stmt(
                        binary(BinOp::In, ident("__k"), this_field("_headers")),
                        vec![ret(index(this_field("_headers"), ident("__k")))],
                    ),
                    ret(ident("default")),
                ],
            ),
            method(
                "items",
                vec![],
                vec![ret(call_global(
                    "list",
                    vec![call(member(this_field("_headers"), "items"), vec![])],
                ))],
            ),
            method(
                "keys",
                vec![],
                vec![ret(call_global(
                    "list",
                    vec![call(member(this_field("_headers"), "keys"), vec![])],
                ))],
            ),
        ],
    )
}

pub(super) fn http_response() -> Statement {
    class(
        "HTTPResponse",
        vec![
            init(
                vec![
                    param("status", Some(num(200.0))),
                    param("reason", Some(str_lit("OK"))),
                    param("body", Some(str_lit(""))),
                ],
                vec![
                    set_this("status", ident("status")),
                    set_this("reason", ident("reason")),
                    set_this("_body", ident("body")),
                    set_this("headers", new("HTTPMessage", vec![])),
                ],
            ),
            stub("read", this_field("_body")),
            method(
                "getheader",
                vec![param("name", None), param("default", Some(null()))],
                vec![ret(call(
                    member(this_field("headers"), "get"),
                    vec![ident("name"), ident("default")],
                ))],
            ),
            method(
                "getheaders",
                vec![],
                vec![ret(call(member(this_field("headers"), "items"), vec![]))],
            ),
            stub("close", null()),
        ],
    )
}

pub(super) fn http_connection() -> Statement {
    class(
        "HTTPConnection",
        vec![
            init(
                vec![
                    param("host", None),
                    param("port", Some(num(80.0))),
                    param("timeout", Some(null())),
                ],
                vec![
                    set_this("host", ident("host")),
                    set_this("port", ident("port")),
                    set_this("timeout", ident("timeout")),
                    set_this("sock", null()),
                    set_this("_response", null()),
                ],
            ),
            stub("connect", null()),
            method(
                "request",
                vec![
                    param("method", None),
                    param("url", None),
                    param("body", Some(null())),
                    param("headers", Some(null())),
                ],
                vec![set_this(
                    "_response",
                    new("HTTPResponse", vec![num(200.0), str_lit("OK"), str_lit("")]),
                )],
            ),
            method(
                "getresponse",
                vec![],
                vec![
                    if_stmt(
                        is_none(this_field("_response")),
                        vec![set_this(
                            "_response",
                            new("HTTPResponse", vec![num(200.0), str_lit("OK"), str_lit("")]),
                        )],
                    ),
                    ret(this_field("_response")),
                ],
            ),
            stub("close", null()),
        ],
    )
}

/// `HTTPSConnection` — `HTTPConnection` with the default port changed. The
/// inherited constructor does the rest, which is the whole point of declaring
/// the parent.
pub(super) fn https_connection() -> Statement {
    class_extending(
        "HTTPSConnection",
        &["HTTPConnection"],
        vec![init(
            vec![
                param("host", None),
                param("port", Some(num(443.0))),
                param("timeout", Some(null())),
            ],
            vec![
                set_this("host", ident("host")),
                set_this("port", ident("port")),
                set_this("timeout", ident("timeout")),
                set_this("sock", null()),
                set_this("_response", null()),
            ],
        )],
    )
}

/// The exception tree. Each parent is a catchability statement: `except
/// HTTPException` catching a `BadStatusLine` IS this declaration.
pub(super) const EXCEPTIONS: &[(&str, &str)] = &[
    ("HTTPException", "Exception"),
    ("BadStatusLine", "HTTPException"),
    ("IncompleteRead", "HTTPException"),
    ("SSLError", "OSError"),
    ("CertificateError", "SSLError"),
];

pub(super) fn exception(name: &'static str, parent: &'static str) -> Statement {
    class_extending(name, &[parent], vec![])
}

pub(super) fn ssl_context() -> Statement {
    class(
        "SSLContext",
        vec![
            init(
                vec![param("protocol", Some(num(2.0)))],
                vec![
                    set_this("protocol", ident("protocol")),
                    set_this("verify_mode", num(0.0)),
                    set_this("check_hostname", bool_lit(false)),
                ],
            ),
            stub("get_ciphers", list_of(vec![])),
            stub("set_ciphers", null()),
            stub("load_verify_locations", null()),
            stub("load_default_certs", null()),
            // `wrap_socket` answers the socket unchanged: there is no TLS
            // layer, and handing back a different object would break every
            // caller that keeps using it.
            method(
                "wrap_socket",
                vec![param("sock", None), rest_param("a"), kwargs_param("k")],
                vec![ret(ident("sock"))],
            ),
        ],
    )
}

/// `ssl.TLSVersion` / `ssl.Purpose` — constant holders. Declared as classes
/// because that is what they are in CPython (`ssl.TLSVersion.TLSv1_2`), and a
/// static field is the declaration for a class-level constant.
pub(super) fn tls_version() -> Statement {
    class(
        "TLSVersion",
        vec![
            static_field("TLSv1", num(769.0)),
            static_field("TLSv1_1", num(770.0)),
            static_field("TLSv1_2", num(771.0)),
            static_field("TLSv1_3", num(772.0)),
        ],
    )
}

pub(super) fn purpose() -> Statement {
    class(
        "Purpose",
        vec![
            static_field("SERVER_AUTH", str_lit("serverAuth")),
            static_field("CLIENT_AUTH", str_lit("clientAuth")),
        ],
    )
}

pub(super) fn module_functions() -> Vec<Statement> {
    vec![
        stub_fn("parse_headers", new("HTTPMessage", vec![])),
        function(
            "create_default_context",
            any_args(),
            vec![ret(new("SSLContext", vec![num(2.0)]))],
        ),
        function(
            "ssl_wrap_socket",
            vec![param("sock", None), rest_param("a"), kwargs_param("k")],
            vec![ret(ident("sock"))],
        ),
        stub_fn("match_hostname", null()),
        stub_fn("enum_certificates", list_of(vec![])),
    ]
}

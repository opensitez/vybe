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

type Expr = vybe_ast::Expression;

fn i(n: i64) -> Expr {
    Expr::int(n)
}

fn op(op: BinOp, left: Expr, right: Expr) -> Expr {
    binary(op, left, right)
}

fn set_attr_this(name: &str, value: Expr) -> Statement {
    expr_stmt(call_global(
        "__py_attr_raw_write",
        vec![ident("self"), str_lit(name), value],
    ))
}

fn status_new(value: i64) -> Expr {
    new("HTTPStatus", vec![i(value)])
}

fn status_case(code: i64, phrase: &str, description: &str) -> Statement {
    if_stmt(
        op(BinOp::Eq, ident("value"), i(code)),
        vec![
            set_this("value", i(code)),
            set_this("phrase", str_lit(phrase)),
            set_this("description", str_lit(description)),
            set_this("is_informational", op(BinOp::Lt, ident("value"), i(200))),
            set_this(
                "is_success",
                op(
                    BinOp::And,
                    op(BinOp::GtEq, ident("value"), i(200)),
                    op(BinOp::Lt, ident("value"), i(300)),
                ),
            ),
            set_this(
                "is_redirection",
                op(
                    BinOp::And,
                    op(BinOp::GtEq, ident("value"), i(300)),
                    op(BinOp::Lt, ident("value"), i(400)),
                ),
            ),
            set_this(
                "is_client_error",
                op(
                    BinOp::And,
                    op(BinOp::GtEq, ident("value"), i(400)),
                    op(BinOp::Lt, ident("value"), i(500)),
                ),
            ),
            set_this("is_server_error", op(BinOp::GtEq, ident("value"), i(500))),
            ret(ident("self")),
        ],
    )
}

fn morsel_attrs() -> &'static [&'static str] {
    &[
        "expires", "path", "comment", "domain", "max-age", "secure", "version", "httponly",
        "samesite",
    ]
}

fn morsel_attr_map() -> Expr {
    dict_str(
        morsel_attrs()
            .iter()
            .map(|name| (*name, this_field(name)))
            .collect(),
    )
}

fn morsel_getitem_body() -> Vec<Statement> {
    let mut body = Vec::new();
    for name in morsel_attrs() {
        body.push(if_stmt(
            op(BinOp::Eq, ident("key"), str_lit(name)),
            vec![ret(this_field(name))],
        ));
    }
    body.push(raise_new("KeyError", vec![ident("key")]));
    body
}

fn morsel_setitem_body() -> Vec<Statement> {
    let mut body = Vec::new();
    for name in morsel_attrs() {
        body.push(if_stmt(
            op(BinOp::Eq, ident("key"), str_lit(name)),
            vec![set_attr_this(name, ident("value")), ret(null())],
        ));
    }
    body.push(raise_new("KeyError", vec![ident("key")]));
    body
}

fn morsel_contains_expr() -> Expr {
    let mut attrs = morsel_attrs().iter();
    let first = attrs
        .next()
        .map(|name| op(BinOp::Eq, ident("key"), str_lit(name)))
        .unwrap_or_else(|| bool_lit(false));
    attrs.fold(first, |acc, name| {
        op(BinOp::Or, acc, op(BinOp::Eq, ident("key"), str_lit(name)))
    })
}

pub(super) fn http_status() -> Statement {
    let mut members = vec![
        static_field("CONTINUE", status_new(100)),
        static_field("SWITCHING_PROTOCOLS", status_new(101)),
        static_field("OK", status_new(200)),
        static_field("CREATED", status_new(201)),
        static_field("NO_CONTENT", status_new(204)),
        static_field("MOVED_PERMANENTLY", status_new(301)),
        static_field("FOUND", status_new(302)),
        static_field("BAD_REQUEST", status_new(400)),
        static_field("UNAUTHORIZED", status_new(401)),
        static_field("FORBIDDEN", status_new(403)),
        static_field("NOT_FOUND", status_new(404)),
        static_field("METHOD_NOT_ALLOWED", status_new(405)),
        static_field("TOO_MANY_REQUESTS", status_new(429)),
        static_field("INTERNAL_SERVER_ERROR", status_new(500)),
        init(
            vec![param("value", None)],
            vec![
                status_case(100, "Continue", "Request received, please continue"),
                status_case(101, "Switching Protocols", "Switching to new protocol"),
                status_case(200, "OK", "Request fulfilled"),
                status_case(201, "Created", "Document created"),
                status_case(204, "No Content", "Request fulfilled, nothing follows"),
                status_case(301, "Moved Permanently", "Object moved permanently"),
                status_case(302, "Found", "Object moved temporarily"),
                status_case(400, "Bad Request", "Bad request syntax"),
                status_case(401, "Unauthorized", "No permission"),
                status_case(403, "Forbidden", "Request forbidden"),
                status_case(404, "Not Found", "Nothing matches the given URI"),
                status_case(405, "Method Not Allowed", "Specified method is invalid"),
                status_case(429, "Too Many Requests", "Too many requests"),
                status_case(500, "Internal Server Error", "Server got itself in trouble"),
                raise_call("ValueError", vec![str_lit("invalid HTTP status")]),
            ],
        ),
        method(
            "__str__",
            vec![],
            vec![ret(call_global("str", vec![this_field("value")]))],
        ),
    ];
    members.push(method(
        "__repr__",
        vec![],
        vec![ret(call_global("str", vec![this_field("value")]))],
    ));
    class("HTTPStatus", members)
}

pub(super) fn morsel() -> Statement {
    class(
        "Morsel",
        vec![
            init(
                vec![
                    param("key", Some(str_lit(""))),
                    param("value", Some(str_lit(""))),
                    param("coded", Some(null())),
                ],
                vec![
                    set_attr_this("key", ident("key")),
                    set_attr_this("value", ident("value")),
                    set_attr_this(
                        "coded_value",
                        ternary(
                            is_none(ident("coded")),
                            ternary(
                                op(
                                    BinOp::GtEq,
                                    call(member(ident("value"), "find"), vec![str_lit(" ")]),
                                    i(0),
                                ),
                                add(add(str_lit("\""), ident("value")), str_lit("\"")),
                                ident("value"),
                            ),
                            ident("coded"),
                        ),
                    ),
                    set_attr_this("expires", null()),
                    set_attr_this("path", null()),
                    set_attr_this("comment", null()),
                    set_attr_this("domain", null()),
                    set_attr_this("max-age", null()),
                    set_attr_this("secure", null()),
                    set_attr_this("version", null()),
                    set_attr_this("httponly", null()),
                    set_attr_this("samesite", null()),
                ],
            ),
            method(
                "set",
                vec![
                    param("key", None),
                    param("value", None),
                    param("coded", None),
                ],
                vec![
                    set_attr_this("key", ident("key")),
                    set_attr_this("value", ident("value")),
                    set_attr_this("coded_value", ident("coded")),
                ],
            ),
            method(
                "__getitem__",
                vec![param("key", None)],
                morsel_getitem_body(),
            ),
            method(
                "__setitem__",
                vec![param("key", None), param("value", None)],
                morsel_setitem_body(),
            ),
            method(
                "__contains__",
                vec![param("key", None)],
                vec![ret(morsel_contains_expr())],
            ),
            method(
                "keys",
                vec![],
                vec![ret(list_of(vec![
                    str_lit("expires"),
                    str_lit("path"),
                    str_lit("comment"),
                    str_lit("domain"),
                    str_lit("max-age"),
                    str_lit("secure"),
                    str_lit("version"),
                    str_lit("httponly"),
                    str_lit("samesite"),
                ]))],
            ),
            method(
                "OutputString",
                vec![],
                vec![
                    assign(
                        ident("__out"),
                        call_global(
                            "__py_cookie_serialize",
                            vec![this_field("key"), this_field("value"), morsel_attr_map()],
                        ),
                    ),
                    if_stmt(
                        is_not_none(this_field("max-age")),
                        vec![assign(
                            ident("__out"),
                            add(
                                add(ident("__out"), str_lit("; Max-Age=")),
                                call_global("str", vec![this_field("max-age")]),
                            ),
                        )],
                    ),
                    ret(ident("__out")),
                ],
            ),
            method(
                "__str__",
                vec![],
                vec![ret(call(member(ident("self"), "OutputString"), vec![]))],
            ),
        ],
    )
}

pub(super) fn simple_cookie() -> Statement {
    class(
        "SimpleCookie",
        vec![
            init(
                vec![param("input", Some(null()))],
                vec![
                    set_attr_this("_cookies", dict_of(vec![])),
                    set_attr_this("_keys", list_of(vec![])),
                    if_stmt(
                        is_not_none(ident("input")),
                        vec![expr_stmt(call(
                            member(ident("self"), "load"),
                            vec![ident("input")],
                        ))],
                    ),
                ],
            ),
            method(
                "__setitem__",
                vec![param("key", None), param("value", None)],
                vec![
                    if_stmt(
                        op(
                            BinOp::GtEq,
                            call(member(ident("key"), "find"), vec![str_lit(" ")]),
                            i(0),
                        ),
                        vec![raise_new(
                            "CookieError",
                            vec![str_lit("invalid cookie key")],
                        )],
                    ),
                    if_stmt(
                        unary_not(call_global(
                            "__py_contains__",
                            vec![this_field("_cookies"), ident("key")],
                        )),
                        vec![expr_stmt(call(
                            member(this_field("_keys"), "append"),
                            vec![ident("key")],
                        ))],
                    ),
                    assign(
                        index(this_field("_cookies"), ident("key")),
                        new(
                            "Morsel",
                            vec![ident("key"), call_global("str", vec![ident("value")])],
                        ),
                    ),
                ],
            ),
            method(
                "__getitem__",
                vec![param("key", None)],
                vec![ret(index(this_field("_cookies"), ident("key")))],
            ),
            method(
                "__len__",
                vec![],
                vec![ret(call_global("len", vec![this_field("_cookies")]))],
            ),
            method("keys", vec![], vec![ret(this_field("_keys"))]),
            method(
                "clear",
                vec![],
                vec![
                    set_attr_this("_cookies", dict_of(vec![])),
                    set_attr_this("_keys", list_of(vec![])),
                ],
            ),
            method(
                "load",
                vec![param("raw", None)],
                vec![
                    if_stmt(
                        unary_not(call_global("hasattr", vec![ident("raw"), str_lit("keys")])),
                        vec![
                            for_in(
                                "__part",
                                call(member(ident("raw"), "split"), vec![str_lit(";")]),
                                vec![
                                    assign(
                                        ident("__part"),
                                        call(member(ident("__part"), "strip"), vec![]),
                                    ),
                                    assign(
                                        ident("__eq"),
                                        call(member(ident("__part"), "find"), vec![str_lit("=")]),
                                    ),
                                    if_stmt(
                                        op(BinOp::GtEq, ident("__eq"), i(0)),
                                        vec![expr_stmt(call(
                                            member(ident("self"), "__setitem__"),
                                            vec![
                                                call(
                                                    member(
                                                        slice_range(
                                                            ident("__part"),
                                                            i(0),
                                                            ident("__eq"),
                                                        ),
                                                        "strip",
                                                    ),
                                                    vec![],
                                                ),
                                                call(
                                                    member(
                                                        slice_from(
                                                            ident("__part"),
                                                            op(BinOp::Add, ident("__eq"), i(1)),
                                                        ),
                                                        "strip",
                                                    ),
                                                    vec![],
                                                ),
                                            ],
                                        ))],
                                    ),
                                ],
                            ),
                            ret(null()),
                        ],
                    ),
                    for_in(
                        "__k",
                        ident("raw"),
                        vec![expr_stmt(call(
                            member(ident("self"), "__setitem__"),
                            vec![ident("__k"), index(ident("raw"), ident("__k"))],
                        ))],
                    ),
                ],
            ),
            method(
                "output",
                vec![
                    param("attrs", Some(null())),
                    param("header", Some(str_lit("Set-Cookie:"))),
                    param("sep", Some(str_lit("\r\n"))),
                ],
                vec![
                    assign(ident("__out"), str_lit("")),
                    for_in(
                        "__k",
                        this_field("_keys"),
                        vec![
                            if_stmt(
                                op(BinOp::NotEq, ident("__out"), str_lit("")),
                                vec![assign(ident("__out"), add(ident("__out"), str_lit("\r\n")))],
                            ),
                            assign(
                                ident("__out"),
                                add(
                                    add(ident("__out"), str_lit("Set-Cookie: ")),
                                    call(
                                        member(
                                            index(this_field("_cookies"), ident("__k")),
                                            "OutputString",
                                        ),
                                        vec![],
                                    ),
                                ),
                            ),
                        ],
                    ),
                    ret(ident("__out")),
                ],
            ),
            method(
                "js_output",
                vec![param("attrs", Some(null()))],
                vec![
                    assign(ident("__out"), str_lit("")),
                    for_in(
                        "__k",
                        this_field("_keys"),
                        vec![
                            if_stmt(
                                op(BinOp::NotEq, ident("__out"), str_lit("")),
                                vec![assign(ident("__out"), add(ident("__out"), str_lit("\r\n")))],
                            ),
                            assign(
                                ident("__out"),
                                add(
                                    ident("__out"),
                                    call(
                                        member(
                                            index(this_field("_cookies"), ident("__k")),
                                            "OutputString",
                                        ),
                                        vec![],
                                    ),
                                ),
                            ),
                        ],
                    ),
                    ret(add(
                        add(
                            str_lit(
                                "<script type=\"text/javascript\">\n        <!-- begin hiding\n        document.cookie = decodeURIComponent(\"",
                            ),
                            call_global("quote", vec![ident("__out")]),
                        ),
                        str_lit("\");\n        // end hiding -->\n        </script>"),
                    )),
                ],
            ),
        ],
    )
}

pub(super) fn cookie_jar() -> Statement {
    class(
        "CookieJar",
        vec![
            init(vec![], vec![set_this("_cookies", list_of(vec![]))]),
            method(
                "__iter__",
                vec![],
                vec![ret(call_global("iter", vec![this_field("_cookies")]))],
            ),
        ],
    )
}

pub(super) fn lwp_cookie_jar() -> Statement {
    class_extending(
        "LWPCookieJar",
        &["CookieJar"],
        vec![
            init(
                vec![param("filename", Some(null()))],
                vec![
                    set_this("_cookies", list_of(vec![])),
                    set_this("filename", ident("filename")),
                ],
            ),
            method("save", any_args(), vec![ret(null())]),
            method("load", any_args(), vec![ret(null())]),
        ],
    )
}

pub(super) fn urllib_request() -> Statement {
    class(
        "Request",
        vec![init(
            vec![
                param("url", None),
                param("data", Some(null())),
                param("headers", Some(null())),
            ],
            vec![
                set_this("full_url", ident("url")),
                set_this("data", ident("data")),
                set_this("headers", dict_of(vec![])),
                if_stmt(
                    is_not_none(ident("headers")),
                    vec![for_in(
                        "__k",
                        call(member(ident("headers"), "keys"), vec![]),
                        vec![assign(
                            index(
                                this_field("headers"),
                                add(
                                    call(
                                        member(slice_range(ident("__k"), i(0), i(1)), "upper"),
                                        vec![],
                                    ),
                                    call(member(slice_from(ident("__k"), i(1)), "lower"), vec![]),
                                ),
                            ),
                            index(ident("headers"), ident("__k")),
                        )],
                    )],
                ),
            ],
        )],
    )
}

/// `http.client.HTTPMessage` — the header bag on a response.
pub(super) fn http_message() -> Statement {
    class(
        "HTTPMessage",
        vec![
            init(
                vec![param("headers", Some(null())), param("keys", Some(null()))],
                vec![
                    set_attr_this(
                        "_headers",
                        ternary(
                            is_none(ident("headers")),
                            call_global("dict", vec![]),
                            ident("headers"),
                        ),
                    ),
                    set_attr_this("_keys", list_of(vec![])),
                    if_stmt(
                        is_not_none(ident("keys")),
                        vec![set_attr_this("_keys", ident("keys"))],
                    ),
                    if_stmt(
                        op(
                            BinOp::And,
                            is_none(ident("keys")),
                            is_not_none(ident("headers")),
                        ),
                        vec![for_in(
                            "__k",
                            ident("headers"),
                            vec![expr_stmt(call(
                                member(this_field("_keys"), "append"),
                                vec![ident("__k")],
                            ))],
                        )],
                    ),
                ],
            ),
            method(
                "get",
                vec![param("name", None), param("default", Some(null()))],
                vec![ret(call_global(
                    "__py_http_message_get",
                    vec![ident("self"), ident("name"), ident("default")],
                ))],
            ),
            method(
                "items",
                vec![],
                vec![ret(call_global(
                    "__py_http_message_items",
                    vec![ident("self")],
                ))],
            ),
            method("keys", vec![], vec![ret(this_field("_keys"))]),
            method(
                "__getitem__",
                vec![param("name", None)],
                vec![ret(call_global(
                    "__py_http_message_get",
                    vec![ident("self"), ident("name"), null()],
                ))],
            ),
            method(
                "get_content_type",
                vec![],
                vec![ret(call_global(
                    "__py_http_message_get_content_type",
                    vec![ident("self")],
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
                    set_this("_sock", ident("status")),
                    set_this("status", num(200.0)),
                    set_this("reason", str_lit("OK")),
                    set_this("_body_text", str_lit("")),
                    set_this("_body_pos", i(0)),
                    set_this("_closed", bool_lit(false)),
                    set_this("version", num(11.0)),
                    set_this("headers", new("HTTPMessage", vec![])),
                ],
            ),
            method(
                "read",
                vec![param("size", Some(i(-1)))],
                vec![
                    assign(
                        ident("__r"),
                        slice_from(this_field("_body_text"), this_field("_body_pos")),
                    ),
                    if_stmt(
                        op(BinOp::GtEq, ident("size"), i(0)),
                        vec![assign(
                            ident("__r"),
                            slice_range(ident("__r"), i(0), ident("size")),
                        )],
                    ),
                    set_this(
                        "_body_pos",
                        op(
                            BinOp::Add,
                            this_field("_body_pos"),
                            call_global("len", vec![ident("__r")]),
                        ),
                    ),
                    ret(call_global("bytes", vec![ident("__r"), str_lit("utf-8")])),
                ],
            ),
            method(
                "readline",
                vec![],
                vec![
                    assign(
                        ident("__rest"),
                        slice_from(this_field("_body_text"), this_field("_body_pos")),
                    ),
                    assign(
                        ident("__i"),
                        call(member(ident("__rest"), "find"), vec![str_lit("\n")]),
                    ),
                    assign(ident("__r"), ident("__rest")),
                    if_stmt(
                        op(BinOp::GtEq, ident("__i"), i(0)),
                        vec![assign(
                            ident("__r"),
                            slice_range(ident("__rest"), i(0), op(BinOp::Add, ident("__i"), i(1))),
                        )],
                    ),
                    set_this(
                        "_body_pos",
                        op(
                            BinOp::Add,
                            this_field("_body_pos"),
                            call_global("len", vec![ident("__r")]),
                        ),
                    ),
                    ret(call_global("bytes", vec![ident("__r"), str_lit("utf-8")])),
                ],
            ),
            method(
                "readinto",
                vec![param("buf", None)],
                vec![
                    assign(
                        ident("__data"),
                        call(
                            member(call(member(ident("self"), "read"), vec![]), "decode"),
                            vec![],
                        ),
                    ),
                    assign(ident("__i"), i(0)),
                    while_stmt(
                        op(
                            BinOp::Lt,
                            ident("__i"),
                            call_global("len", vec![ident("__data")]),
                        ),
                        vec![
                            assign(
                                index(ident("buf"), ident("__i")),
                                call_global("ord", vec![index(ident("__data"), ident("__i"))]),
                            ),
                            assign(ident("__i"), op(BinOp::Add, ident("__i"), i(1))),
                        ],
                    ),
                    ret(call_global("len", vec![ident("__data")])),
                ],
            ),
            method(
                "begin",
                vec![],
                vec![
                    if_stmt(is_none(this_field("_sock")), vec![ret(null())]),
                    assign(
                        ident("__text"),
                        call_global("__py_http_socket_text", vec![this_field("_sock")]),
                    ),
                    assign(
                        ident("__text"),
                        call(
                            member(ident("__text"), "replace"),
                            vec![str_lit("\r\n"), str_lit("\n")],
                        ),
                    ),
                    assign(
                        ident("__split"),
                        call(member(ident("__text"), "find"), vec![str_lit("\n\n")]),
                    ),
                    assign(ident("__sep"), i(2)),
                    assign(
                        ident("__head"),
                        slice_range(ident("__text"), i(0), ident("__split")),
                    ),
                    assign(
                        ident("__body"),
                        slice_from(
                            ident("__text"),
                            op(BinOp::Add, ident("__split"), ident("__sep")),
                        ),
                    ),
                    assign(
                        ident("__line_end"),
                        call(member(ident("__head"), "find"), vec![str_lit("\n")]),
                    ),
                    assign(ident("__status_line"), ident("__head")),
                    assign(ident("__header_text"), str_lit("")),
                    if_stmt(
                        op(BinOp::GtEq, ident("__line_end"), i(0)),
                        vec![
                            assign(
                                ident("__status_line"),
                                slice_range(ident("__head"), i(0), ident("__line_end")),
                            ),
                            assign(
                                ident("__header_text"),
                                slice_from(
                                    ident("__head"),
                                    op(BinOp::Add, ident("__line_end"), i(1)),
                                ),
                            ),
                        ],
                    ),
                    assign(
                        ident("__status"),
                        call(
                            member(ident("__status_line"), "split"),
                            vec![str_lit(" "), i(2)],
                        ),
                    ),
                    set_attr_this(
                        "version",
                        ternary(
                            op(
                                BinOp::Eq,
                                index(ident("__status"), i(0)),
                                str_lit("HTTP/1.1"),
                            ),
                            i(11),
                            i(10),
                        ),
                    ),
                    set_attr_this(
                        "status",
                        call_global("int", vec![index(ident("__status"), i(1))]),
                    ),
                    set_attr_this(
                        "reason",
                        ternary(
                            op(BinOp::Gt, call_global("len", vec![ident("__status")]), i(2)),
                            index(ident("__status"), i(2)),
                            str_lit(""),
                        ),
                    ),
                    set_attr_this(
                        "headers",
                        call_global("__py_http_parse_headers_text", vec![ident("__header_text")]),
                    ),
                    assign(
                        ident("__headers"),
                        field_of(this_field("headers"), "_headers"),
                    ),
                    if_stmt(
                        op(
                            BinOp::And,
                            call_global(
                                "__py_contains__",
                                vec![ident("__headers"), str_lit("transfer-encoding")],
                            ),
                            op(
                                BinOp::Eq,
                                call(
                                    member(
                                        call_global(
                                            "str",
                                            vec![index(
                                                ident("__headers"),
                                                str_lit("transfer-encoding"),
                                            )],
                                        ),
                                        "lower",
                                    ),
                                    vec![],
                                ),
                                str_lit("chunked"),
                            ),
                        ),
                        vec![assign(
                            ident("__body"),
                            call_global("__py_http_decode_chunked", vec![ident("__body")]),
                        )],
                    ),
                    set_attr_this("_body_text", ident("__body")),
                    set_attr_this("_body_pos", i(0)),
                    ret(null()),
                ],
            ),
            method(
                "getheader",
                vec![param("name", None), param("default", Some(null()))],
                vec![
                    assign(
                        ident("__headers"),
                        field_of(this_field("headers"), "_headers"),
                    ),
                    assign(
                        ident("__k"),
                        call(
                            member(call_global("str", vec![ident("name")]), "lower"),
                            vec![],
                        ),
                    ),
                    if_stmt(
                        call_global("__py_contains__", vec![ident("__headers"), ident("__k")]),
                        vec![ret(index(ident("__headers"), ident("__k")))],
                    ),
                    ret(ident("default")),
                ],
            ),
            method(
                "getheaders",
                vec![],
                vec![ret(call_global(
                    "__py_http_message_items",
                    vec![this_field("headers")],
                ))],
            ),
            method("isclosed", vec![], vec![ret(this_field("_closed"))]),
            method("close", vec![], vec![set_this("_closed", bool_lit(true))]),
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
                    set_this("_tunnel_host", null()),
                    set_this("_tunnel_port", null()),
                    set_this("_out", str_lit("")),
                    set_this("_HTTPConnection__state", i(0)),
                ],
            ),
            stub("connect", null()),
            method(
                "putrequest",
                vec![param("method", None), param("url", None)],
                vec![
                    if_stmt(
                        op(BinOp::Eq, this_field("_HTTPConnection__state"), i(2)),
                        vec![raise_new("CannotSendRequest", vec![])],
                    ),
                    set_this(
                        "_out",
                        add(
                            add(add(ident("method"), str_lit(" ")), ident("url")),
                            str_lit(" HTTP/1.1\r\n"),
                        ),
                    ),
                    set_this("_HTTPConnection__state", i(2)),
                ],
            ),
            method(
                "putheader",
                vec![param("header", None), rest_param("values")],
                vec![set_this(
                    "_out",
                    add(
                        add(
                            add(add(this_field("_out"), ident("header")), str_lit(": ")),
                            call_global("str", vec![index(ident("values"), i(0))]),
                        ),
                        str_lit("\r\n"),
                    ),
                )],
            ),
            method(
                "endheaders",
                vec![],
                vec![if_stmt(
                    is_not_none(this_field("sock")),
                    vec![expr_stmt(call_global(
                        "__py_http_send",
                        vec![this_field("sock"), add(this_field("_out"), str_lit("\r\n"))],
                    ))],
                )],
            ),
            method(
                "set_tunnel",
                vec![
                    param("host", None),
                    param("port", Some(null())),
                    param("headers", Some(null())),
                ],
                vec![
                    set_this("_tunnel_host", ident("host")),
                    set_this("_tunnel_port", ident("port")),
                ],
            ),
            method(
                "request",
                vec![
                    param("method", None),
                    param("url", None),
                    param("body", Some(null())),
                    param("headers", Some(null())),
                ],
                vec![
                    expr_stmt(call(
                        member(ident("self"), "putrequest"),
                        vec![ident("method"), ident("url")],
                    )),
                    if_stmt(
                        is_not_none(ident("headers")),
                        vec![for_in(
                            "__k",
                            call(member(ident("headers"), "keys"), vec![]),
                            vec![expr_stmt(call(
                                member(ident("self"), "putheader"),
                                vec![ident("__k"), index(ident("headers"), ident("__k"))],
                            ))],
                        )],
                    ),
                    if_stmt(
                        is_not_none(ident("body")),
                        vec![expr_stmt(call(
                            member(ident("self"), "putheader"),
                            vec![
                                str_lit("Content-Length"),
                                call_global("str", vec![call_global("len", vec![ident("body")])]),
                            ],
                        ))],
                    ),
                    expr_stmt(call(member(ident("self"), "endheaders"), vec![])),
                    if_stmt(
                        op(
                            BinOp::And,
                            is_not_none(this_field("sock")),
                            is_not_none(ident("body")),
                        ),
                        vec![expr_stmt(call_global(
                            "__py_http_send",
                            vec![this_field("sock"), ident("body")],
                        ))],
                    ),
                    set_this(
                        "_response",
                        new("HTTPResponse", vec![num(200.0), str_lit("OK"), str_lit("")]),
                    ),
                ],
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
    ("CannotSendRequest", "HTTPException"),
    ("CookieError", "Exception"),
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
        global_assign(
            "responses",
            dict_of(vec![
                (i(100), str_lit("Continue")),
                (i(101), str_lit("Switching Protocols")),
                (i(200), str_lit("OK")),
                (i(201), str_lit("Created")),
                (i(204), str_lit("No Content")),
                (i(301), str_lit("Moved Permanently")),
                (i(302), str_lit("Found")),
                (i(400), str_lit("Bad Request")),
                (i(401), str_lit("Unauthorized")),
                (i(403), str_lit("Forbidden")),
                (i(404), str_lit("Not Found")),
                (i(405), str_lit("Method Not Allowed")),
                (i(429), str_lit("Too Many Requests")),
                (i(500), str_lit("Internal Server Error")),
            ]),
        ),
        function(
            "__py_http_send",
            vec![param("sock", None), param("data", None)],
            vec![
                if_stmt(
                    call_global("hasattr", vec![ident("sock"), str_lit("buf")]),
                    vec![
                        assign(ident("__http_buf"), field_of(ident("sock"), "buf")),
                        assign(
                            ident("__text"),
                            call_global("__py_io_text", vec![ident("data")]),
                        ),
                        expr_stmt(call(
                            member(ident("__http_buf"), "write"),
                            vec![ident("__text")],
                        )),
                        if_stmt(
                            op(
                                BinOp::Eq,
                                call_global("len", vec![field_of(ident("__http_buf"), "_buf")]),
                                i(0),
                            ),
                            vec![
                                expr_stmt(call_global(
                                    "__py_attr_raw_write",
                                    vec![
                                        ident("__http_buf"),
                                        str_lit("_buf"),
                                        add(field_of(ident("__http_buf"), "_buf"), ident("__text")),
                                    ],
                                )),
                                expr_stmt(call_global(
                                    "__py_attr_raw_write",
                                    vec![
                                        ident("__http_buf"),
                                        str_lit("_pos"),
                                        call_global(
                                            "len",
                                            vec![field_of(ident("__http_buf"), "_buf")],
                                        ),
                                    ],
                                )),
                            ],
                        ),
                        ret(null()),
                    ],
                ),
                if_stmt(
                    call_global("hasattr", vec![ident("sock"), str_lit("sendall")]),
                    vec![expr_stmt(call(
                        member(ident("sock"), "sendall"),
                        vec![ident("data")],
                    ))],
                ),
                ret(null()),
            ],
        ),
        function(
            "__py_http_file_text",
            vec![param("fp", None)],
            vec![
                assign(ident("__raw"), null()),
                if_stmt(
                    call_global("hasattr", vec![ident("fp"), str_lit("_buf")]),
                    vec![assign(ident("__raw"), field_of(ident("fp"), "_buf"))],
                ),
                if_stmt(
                    op(
                        BinOp::And,
                        is_none(ident("__raw")),
                        call_global("hasattr", vec![ident("fp"), str_lit("getvalue")]),
                    ),
                    vec![assign(
                        ident("__raw"),
                        call(member(ident("fp"), "getvalue"), vec![]),
                    )],
                ),
                if_stmt(
                    call_global("hasattr", vec![ident("__raw"), str_lit("decode")]),
                    vec![ret(call(member(ident("__raw"), "decode"), vec![]))],
                ),
                ret(call_global("str", vec![ident("__raw")])),
            ],
        ),
        function(
            "__py_http_socket_text",
            vec![param("sock", None)],
            vec![
                assign(ident("__fp"), null()),
                if_stmt(
                    call_global("hasattr", vec![ident("sock"), str_lit("makefile")]),
                    vec![assign(
                        ident("__fp"),
                        call(member(ident("sock"), "makefile"), vec![str_lit("rb")]),
                    )],
                ),
                if_stmt(
                    is_none(ident("__fp")),
                    vec![assign(ident("__fp"), field_of(ident("sock"), "file"))],
                ),
                ret(call_global("__py_http_file_text", vec![ident("__fp")])),
            ],
        ),
        function(
            "__py_http_parse_headers_text",
            vec![param("text", Some(str_lit("")))],
            vec![
                assign(
                    ident("text"),
                    call(
                        member(ident("text"), "replace"),
                        vec![str_lit("\r\n"), str_lit("\n")],
                    ),
                ),
                assign(ident("__headers"), dict_of(vec![])),
                assign(ident("__keys"), list_of(vec![])),
                for_in(
                    "__line",
                    call(member(ident("text"), "split"), vec![str_lit("\n")]),
                    vec![
                        assign(
                            ident("__line"),
                            call(member(ident("__line"), "strip"), vec![]),
                        ),
                        assign(
                            ident("__pos"),
                            call(member(ident("__line"), "find"), vec![str_lit(":")]),
                        ),
                        if_stmt(
                            op(BinOp::GtEq, ident("__pos"), i(0)),
                            vec![
                                assign(
                                    ident("__name"),
                                    call(
                                        member(
                                            slice_range(ident("__line"), i(0), ident("__pos")),
                                            "strip",
                                        ),
                                        vec![],
                                    ),
                                ),
                                assign(
                                    ident("__value"),
                                    call(
                                        member(
                                            slice_from(
                                                ident("__line"),
                                                op(BinOp::Add, ident("__pos"), i(1)),
                                            ),
                                            "strip",
                                        ),
                                        vec![],
                                    ),
                                ),
                                expr_stmt(call(
                                    member(ident("__keys"), "append"),
                                    vec![ident("__name")],
                                )),
                                assign(
                                    index(ident("__headers"), ident("__name")),
                                    ident("__value"),
                                ),
                                assign(
                                    index(
                                        ident("__headers"),
                                        call(member(ident("__name"), "lower"), vec![]),
                                    ),
                                    ident("__value"),
                                ),
                            ],
                        ),
                    ],
                ),
                ret(new(
                    "HTTPMessage",
                    vec![ident("__headers"), ident("__keys")],
                )),
            ],
        ),
        function(
            "__py_http_message_get",
            vec![
                param("msg", None),
                param("name", None),
                param("default", Some(null())),
            ],
            vec![
                assign(ident("__headers"), field_of(ident("msg"), "_headers")),
                assign(
                    ident("__k"),
                    call(
                        member(call_global("str", vec![ident("name")]), "lower"),
                        vec![],
                    ),
                ),
                if_stmt(
                    call_global("__py_contains__", vec![ident("__headers"), ident("__k")]),
                    vec![ret(index(ident("__headers"), ident("__k")))],
                ),
                ret(ident("default")),
            ],
        ),
        function(
            "__py_http_message_items",
            vec![param("msg", None)],
            vec![
                assign(ident("__out"), list_of(vec![])),
                assign(ident("__headers"), field_of(ident("msg"), "_headers")),
                for_in(
                    "__k",
                    field_of(ident("msg"), "_keys"),
                    vec![expr_stmt(call(
                        member(ident("__out"), "append"),
                        vec![tuple_of(vec![
                            ident("__k"),
                            index(ident("__headers"), ident("__k")),
                        ])],
                    ))],
                ),
                ret(ident("__out")),
            ],
        ),
        function(
            "__py_http_message_get_content_type",
            vec![param("msg", None)],
            vec![ret(call_global(
                "__py_http_message_get",
                vec![ident("msg"), str_lit("content-type"), str_lit("text/plain")],
            ))],
        ),
        function(
            "__py_http_decode_chunked",
            vec![param("body", Some(str_lit("")))],
            vec![
                assign(
                    ident("body"),
                    call(
                        member(ident("body"), "replace"),
                        vec![str_lit("\r\n"), str_lit("\n")],
                    ),
                ),
                assign(ident("__out"), str_lit("")),
                assign(ident("__rest"), ident("body")),
                while_stmt(
                    op(BinOp::Gt, call_global("len", vec![ident("__rest")]), i(0)),
                    vec![
                        assign(
                            ident("__nl"),
                            call(member(ident("__rest"), "find"), vec![str_lit("\n")]),
                        ),
                        if_stmt(
                            op(BinOp::Lt, ident("__nl"), i(0)),
                            vec![ret(ident("__out"))],
                        ),
                        assign(
                            ident("__size_s"),
                            slice_range(ident("__rest"), i(0), ident("__nl")),
                        ),
                        assign(
                            ident("__size"),
                            call_global("int", vec![ident("__size_s"), i(16)]),
                        ),
                        if_stmt(
                            op(BinOp::Eq, ident("__size"), i(0)),
                            vec![ret(ident("__out"))],
                        ),
                        assign(ident("__start"), op(BinOp::Add, ident("__nl"), i(1))),
                        assign(
                            ident("__end"),
                            op(BinOp::Add, ident("__start"), ident("__size")),
                        ),
                        assign(
                            ident("__out"),
                            add(
                                ident("__out"),
                                slice_range(ident("__rest"), ident("__start"), ident("__end")),
                            ),
                        ),
                        assign(
                            ident("__rest"),
                            slice_from(ident("__rest"), op(BinOp::Add, ident("__end"), i(1))),
                        ),
                    ],
                ),
                ret(ident("__out")),
            ],
        ),
        function(
            "parse_headers",
            vec![param("fp", None)],
            vec![
                assign(
                    ident("__text"),
                    call_global("__py_http_file_text", vec![ident("fp")]),
                ),
                ret(call_global(
                    "__py_http_parse_headers_text",
                    vec![ident("__text")],
                )),
            ],
        ),
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

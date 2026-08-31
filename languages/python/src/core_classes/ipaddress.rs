//! `ipaddress` — `IPv4Address`, `IPv4Network`, `IPv4Interface`, as CLASSES.
//!
//! These were 147 lines of `IPADDRESS_PRELUDE` — parsed Python source spliced
//! into every program whose text contained `import ipaddress` — and they are
//! the same classes here, declared as AST. The difference is the whole point:
//! constructed AST needs no second parse, and it flows through
//! `normalize_class` → `NormalClass` → `compile_class`, which is what gives a
//! real rtt, the `TypeRegistry` registration, receiver-based member dispatch
//! and protocol-slot binding.
//!
//! ⛔ The bodies are deliberately PLAIN — arithmetic, `str()`, `+`, `while`.
//! Every one lowers through the shared machinery, so this file carries no
//! python-private emitter and the semantics are the ones every language gets.
//! The address ARITHMETIC that is genuinely bit-twiddling (`_vybe_ip4_parse`,
//! `_vybe_ip4_str`, `_vybe_ip4_mask`, `_vybe_ip4_count`, `_vybe_ip4_octets`,
//! `_vybe_ip4_net_parts`) stays in `emitter/socket_adapter.rs` behind its
//! existing profile rows — that is an adapter's job, and a class calling one is
//! an ordinary call.

use super::builders::*;
use vybe_ast::{BinOp, Statement};

/// `__o[i]` — the octets, bound to a local in `__init__` so each read is a
/// plain index rather than a re-parse.
fn octet(index: f64) -> vybe_ast::Expression {
    index_of(ident("__o"), index)
}

fn index_of(object: vybe_ast::Expression, at: f64) -> vybe_ast::Expression {
    index(object, num(at))
}

fn lt(left: vybe_ast::Expression, right: vybe_ast::Expression) -> vybe_ast::Expression {
    binary(BinOp::Lt, left, right)
}

fn le(left: vybe_ast::Expression, right: vybe_ast::Expression) -> vybe_ast::Expression {
    binary(BinOp::LtEq, left, right)
}

fn ge(left: vybe_ast::Expression, right: vybe_ast::Expression) -> vybe_ast::Expression {
    binary(BinOp::GtEq, left, right)
}

fn eq(left: vybe_ast::Expression, right: vybe_ast::Expression) -> vybe_ast::Expression {
    binary(BinOp::Eq, left, right)
}

fn and(left: vybe_ast::Expression, right: vybe_ast::Expression) -> vybe_ast::Expression {
    binary(BinOp::And, left, right)
}

fn or(left: vybe_ast::Expression, right: vybe_ast::Expression) -> vybe_ast::Expression {
    binary(BinOp::Or, left, right)
}

/// `IPv4Address`.
///
/// `__str__`, `__repr__`, `__int__`, `__eq__`, `__add__` and `__sub__` are
/// declared as ordinary dunder methods. python's `protocol.rs` already maps
/// every one of them onto a `ProtocolSlot`, and `normalize_class` routes them
/// into `NormalClass.special_methods`, so `str(a)`, `int(a)`, `a == b` and
/// `a + 1` bind through the shared slot machinery with nothing stamped by hand.
pub(super) fn ipv4_address() -> Statement {
    class(
        "IPv4Address",
        vec![
            init(
                vec![param("value", None)],
                vec![
                    set_this("version", num(4.0)),
                    set_this("_int", ident("value")),
                    set_this(
                        "_text",
                        call_global("_vybe_ip4_str", vec![ident("value")]),
                    ),
                    set_this("compressed", this_field("_text")),
                    set_this("exploded", this_field("_text")),
                    // ⛔ EAGER FIELDS, not `@property`. Every one is a pure
                    // function of the address, so computing it once at
                    // construction is observationally identical — and
                    // `normalize_class` builds a Property getter with NO `self`
                    // parameter while python is `explicit_self_param`, so a
                    // getter body naming `self` has no receiver to bind.
                    assign(
                        ident("__o"),
                        call_global("_vybe_ip4_octets", vec![ident("value")]),
                    ),
                    set_this("packed", call_global("bytes", vec![ident("__o")])),
                    // 10/8, 127/8, 192.168/16, 172.16/12 — CPython's private ranges.
                    set_this(
                        "is_private",
                        or(
                            or(eq(octet(0.0), num(10.0)), eq(octet(0.0), num(127.0))),
                            or(
                                and(eq(octet(0.0), num(192.0)), eq(octet(1.0), num(168.0))),
                                and(
                                    eq(octet(0.0), num(172.0)),
                                    and(ge(octet(1.0), num(16.0)), le(octet(1.0), num(31.0))),
                                ),
                            ),
                        ),
                    ),
                    set_this("is_loopback", eq(octet(0.0), num(127.0))),
                    set_this(
                        "is_multicast",
                        and(ge(octet(0.0), num(224.0)), le(octet(0.0), num(239.0))),
                    ),
                    set_this(
                        "is_global",
                        eq(this_field("is_private"), bool_lit(false)),
                    ),
                ],
            ),
            method("__str__", vec![], vec![ret(this_field("_text"))]),
            method(
                "__repr__",
                vec![],
                vec![ret(add(
                    add(str_lit("IPv4Address('"), this_field("_text")),
                    str_lit("')"),
                ))],
            ),
            method("__int__", vec![], vec![ret(this_field("_int"))]),
            method(
                "__eq__",
                vec![param("other", None)],
                vec![ret(eq(
                    this_field("_int"),
                    call_global("int", vec![ident("other")]),
                ))],
            ),
            method(
                "__add__",
                vec![param("n", None)],
                vec![ret(new(
                    "IPv4Address",
                    vec![add(this_field("_int"), ident("n"))],
                ))],
            ),
            method(
                "__sub__",
                vec![param("n", None)],
                vec![ret(new(
                    "IPv4Address",
                    vec![binary(BinOp::Sub, this_field("_int"), ident("n"))],
                ))],
            ),
        ],
    )
}

/// `IPv4Network`. `hosts`, `subnets`, `supernet` and `overlaps` are ordinary
/// METHODS — nothing rewrites them, nothing registers them by hand; they
/// dispatch by receiver through the prototype `compile_class` stamps.
pub(super) fn ipv4_network() -> Statement {
    class(
        "IPv4Network",
        vec![
            init(
                vec![param("value", None), param("strict", Some(bool_lit(true)))],
                vec![
                    assign(
                        ident("__pair"),
                        call_global(
                            "_vybe_ip4_net_parts",
                            vec![call_global("str", vec![ident("value")])],
                        ),
                    ),
                    set_this("version", num(4.0)),
                    set_this("prefixlen", index_of(ident("__pair"), 1.0)),
                    set_this(
                        "num_addresses",
                        call_global("_vybe_ip4_count", vec![this_field("prefixlen")]),
                    ),
                    // The network address: the host bits masked off.
                    set_this(
                        "_base",
                        binary(
                            BinOp::Mul,
                            call_global(
                                "int",
                                vec![binary(
                                    BinOp::Div,
                                    index_of(ident("__pair"), 0.0),
                                    this_field("num_addresses"),
                                )],
                            ),
                            this_field("num_addresses"),
                        ),
                    ),
                    // ⛔ `strict=True` (CPython's default) REJECTS an address
                    // with host bits set — `ip_network('192.168.1.5/24')` is a
                    // ValueError, and only `strict=False` truncates. The
                    // prelude ignored the flag entirely.
                    if_stmt(
                        and(
                            ident("strict"),
                            binary(
                                BinOp::NotEq,
                                this_field("_base"),
                                index_of(ident("__pair"), 0.0),
                            ),
                        ),
                        vec![expr_stmt(call_global(
                            "__vybe_raise_value_error",
                            vec![add(
                                call_global("str", vec![ident("value")]),
                                str_lit(" has host bits set"),
                            )],
                        ))],
                    ),
                    assign(
                        ident("__mask"),
                        call_global("_vybe_ip4_mask", vec![this_field("prefixlen")]),
                    ),
                    set_this("network_address", new("IPv4Address", vec![this_field("_base")])),
                    set_this("netmask", new("IPv4Address", vec![ident("__mask")])),
                    set_this(
                        "hostmask",
                        new(
                            "IPv4Address",
                            vec![binary(BinOp::Sub, num(4294967295.0), ident("__mask"))],
                        ),
                    ),
                    set_this(
                        "broadcast_address",
                        new(
                            "IPv4Address",
                            vec![binary(
                                BinOp::Sub,
                                add(this_field("_base"), this_field("num_addresses")),
                                num(1.0),
                            )],
                        ),
                    ),
                ],
            ),
            method(
                "__str__",
                vec![],
                vec![ret(add(
                    add(
                        call_global("str", vec![this_field("network_address")]),
                        str_lit("/"),
                    ),
                    call_global("str", vec![this_field("prefixlen")]),
                ))],
            ),
            method(
                "__repr__",
                vec![],
                vec![ret(add(
                    add(
                        str_lit("IPv4Network('"),
                        call_global("str", vec![ident("self")]),
                    ),
                    str_lit("')"),
                ))],
            ),
            method(
                "__contains__",
                vec![param("addr", None)],
                vec![
                    assign(ident("__v"), call_global("int", vec![ident("addr")])),
                    ret(and(
                        ge(ident("__v"), this_field("_base")),
                        lt(
                            ident("__v"),
                            add(this_field("_base"), this_field("num_addresses")),
                        ),
                    )),
                ],
            ),
            method(
                "hosts",
                vec![],
                vec![
                    assign(ident("__out"), call_global("list", vec![])),
                    assign(ident("__i"), add(this_field("_base"), num(1.0))),
                    assign(
                        ident("__last"),
                        binary(
                            BinOp::Sub,
                            add(this_field("_base"), this_field("num_addresses")),
                            num(1.0),
                        ),
                    ),
                    while_stmt(
                        lt(ident("__i"), ident("__last")),
                        vec![
                            expr_stmt(call(
                                member(ident("__out"), "append"),
                                vec![new("IPv4Address", vec![ident("__i")])],
                            )),
                            assign(ident("__i"), add(ident("__i"), num(1.0))),
                        ],
                    ),
                    ret(ident("__out")),
                ],
            ),
            method(
                "subnets",
                vec![param("prefixlen_diff", Some(num(1.0)))],
                vec![
                    assign(
                        ident("__new_len"),
                        add(this_field("prefixlen"), ident("prefixlen_diff")),
                    ),
                    assign(
                        ident("__step"),
                        call_global("_vybe_ip4_count", vec![ident("__new_len")]),
                    ),
                    assign(ident("__out"), call_global("list", vec![])),
                    assign(ident("__i"), this_field("_base")),
                    assign(
                        ident("__limit"),
                        add(this_field("_base"), this_field("num_addresses")),
                    ),
                    while_stmt(
                        lt(ident("__i"), ident("__limit")),
                        vec![
                            expr_stmt(call(
                                member(ident("__out"), "append"),
                                vec![new(
                                    "IPv4Network",
                                    vec![add(
                                        add(
                                            call_global("_vybe_ip4_str", vec![ident("__i")]),
                                            str_lit("/"),
                                        ),
                                        call_global("str", vec![ident("__new_len")]),
                                    )],
                                )],
                            )),
                            assign(ident("__i"), add(ident("__i"), ident("__step"))),
                        ],
                    ),
                    ret(ident("__out")),
                ],
            ),
            method(
                "supernet",
                vec![param("prefixlen_diff", Some(num(1.0)))],
                vec![ret(new(
                    "IPv4Network",
                    vec![add(
                        add(
                            call_global("_vybe_ip4_str", vec![this_field("_base")]),
                            str_lit("/"),
                        ),
                        call_global(
                            "str",
                            vec![binary(
                                BinOp::Sub,
                                this_field("prefixlen"),
                                ident("prefixlen_diff"),
                            )],
                        ),
                    )],
                ))],
            ),
            method(
                "overlaps",
                vec![param("other", None)],
                vec![ret(and(
                    lt(
                        this_field("_base"),
                        add(
                            field_of(ident("other"), "_base"),
                            field_of(ident("other"), "num_addresses"),
                        ),
                    ),
                    lt(
                        field_of(ident("other"), "_base"),
                        add(this_field("_base"), this_field("num_addresses")),
                    ),
                ))],
            ),
        ],
    )
}

/// `IPv4Interface` — an address that remembers the network it sits in.
pub(super) fn ipv4_interface() -> Statement {
    class(
        "IPv4Interface",
        vec![
            init(
                vec![param("value", None)],
                vec![
                    assign(
                        ident("__pair"),
                        call_global(
                            "_vybe_ip4_net_parts",
                            vec![call_global("str", vec![ident("value")])],
                        ),
                    ),
                    set_this("version", num(4.0)),
                    set_this("prefixlen", index_of(ident("__pair"), 1.0)),
                    set_this("ip", new("IPv4Address", vec![index_of(ident("__pair"), 0.0)])),
                    set_this(
                        "network",
                        new("IPv4Network", vec![ident("value"), bool_lit(false)]),
                    ),
                ],
            ),
            method(
                "__str__",
                vec![],
                vec![ret(add(
                    add(call_global("str", vec![this_field("ip")]), str_lit("/")),
                    call_global("str", vec![this_field("prefixlen")]),
                ))],
            ),
        ],
    )
}

/// The module-level functions, as ordinary declared functions. `ip_network`
/// and `ip_interface` are not here: they ARE the classes, and
/// `MODULE_SURFACE` maps the factory spelling straight onto the class name, so
/// `ipaddress.ip_network(x)` constructs `IPv4Network(x)` with no wrapper.
pub(super) fn module_functions() -> Vec<Statement> {
    vec![
        function(
            "ip_address",
            vec![param("value", None)],
            vec![
                // An int (or packed bytes, which fold to one) is the value
                // directly; text is parsed.
                if_stmt(
                    call_global("isinstance", vec![ident("value"), ident("int")]),
                    vec![ret(new("IPv4Address", vec![ident("value")]))],
                ),
                assign(
                    ident("__v"),
                    call_global(
                        "_vybe_ip4_parse",
                        vec![call_global("str", vec![ident("value")])],
                    ),
                ),
                // ⛔ `ip_address('999.999.999.999')` is a ValueError in
                // CPython. Every octet over 255 pushes the folded value past
                // 2^32, so the range test catches the whole class.
                if_stmt(
                    or(
                        ge(ident("__v"), num(4294967296.0)),
                        lt(ident("__v"), num(0.0)),
                    ),
                    vec![raise_value_error(add(
                        call_global("str", vec![ident("value")]),
                        str_lit(" does not appear to be an IPv4 or IPv6 address"),
                    ))],
                ),
                ret(new("IPv4Address", vec![ident("__v")])),
            ],
        ),
        // Adjacent equal-size networks merge into their supernet. One pass is
        // enough for the pairs CPython's own doctest shows; a full collapse
        // would loop until stable.
        function(
            "collapse_addresses",
            vec![param("nets", None)],
            vec![
                assign(ident("__out"), call_global("list", vec![])),
                assign(ident("__i"), num(0.0)),
                assign(ident("__n"), call_global("len", vec![ident("nets")])),
                while_stmt(
                    lt(ident("__i"), ident("__n")),
                    vec![
                        assign(ident("__a"), index(ident("nets"), ident("__i"))),
                        assign(ident("__merged"), bool_lit(false)),
                        if_stmt(
                            lt(add(ident("__i"), num(1.0)), ident("__n")),
                            vec![
                                assign(
                                    ident("__b"),
                                    index(ident("nets"), add(ident("__i"), num(1.0))),
                                ),
                                if_stmt(
                                    and(
                                        eq(
                                            field_of(ident("__a"), "prefixlen"),
                                            field_of(ident("__b"), "prefixlen"),
                                        ),
                                        eq(
                                            add(
                                                field_of(ident("__a"), "_base"),
                                                field_of(ident("__a"), "num_addresses"),
                                            ),
                                            field_of(ident("__b"), "_base"),
                                        ),
                                    ),
                                    vec![
                                        expr_stmt(call(
                                            member(ident("__out"), "append"),
                                            vec![call(
                                                member(ident("__a"), "supernet"),
                                                vec![],
                                            )],
                                        )),
                                        assign(ident("__merged"), bool_lit(true)),
                                        assign(
                                            ident("__i"),
                                            add(ident("__i"), num(2.0)),
                                        ),
                                    ],
                                ),
                            ],
                        ),
                        if_stmt(
                            eq(ident("__merged"), bool_lit(false)),
                            vec![
                                expr_stmt(call(
                                    member(ident("__out"), "append"),
                                    vec![ident("__a")],
                                )),
                                assign(ident("__i"), add(ident("__i"), num(1.0))),
                            ],
                        ),
                    ],
                ),
                ret(ident("__out")),
            ],
        ),
    ]
}

/// `raise ValueError(<message>)`, as the AST a source `raise` produces.
fn raise_value_error(message: vybe_ast::Expression) -> Statement {
    vybe_ast::Statement::with_span(
        vybe_ast::StmtKind::Throw {
            // ⛔ A CALL, not `New`. `raise ValueError(x)` from source walks to
            // `Throw { Call { Ident("ValueError"), … } }` — the exception
            // classes are not in `py_defined_classes`, so nothing normalises
            // them to `New`, and emitting `New` here gave "undefined is not
            // callable".
            expr: Some(call_global("ValueError", vec![message])),
            cause: None,
        },
        span(),
    )
}

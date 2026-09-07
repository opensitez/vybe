//! `configparser.ConfigParser` — declared, not parsed.
//!
//! Parsing itself stays in `__py_config_parse`, the shared config primitive;
//! this is only the Python object/policy adapter wrapped around it.

use super::builders::*;
use vybe_ast::{BinOp, Statement};

type Expr = vybe_ast::Expression;

const MISSING: &str = "__py_configparser_missing__";

fn sections() -> Expr {
    this_field("_sections")
}

fn defaults_map() -> Expr {
    this_field("_defaults")
}

fn optionxform() -> Expr {
    this_field("_optionxform")
}

/// `x in y`. ⛔ NOT `BinOp::In` — that is a WALKER rewrite, and a spliced
/// declaration is never walked, so it answers False for everything.
fn contains(haystack: Expr, needle: Expr) -> Expr {
    call_global("__py_contains__", vec![haystack, needle])
}

fn sec_of(sec: Expr) -> Expr {
    index(sections(), sec)
}

fn empty_dict() -> Expr {
    Expression::with_span(
        vybe_ast::ExprKind::Map(Vec::new()),
        vybe_ast::Span::default(),
    )
}

use vybe_ast::Expression;

pub(super) const EXCEPTIONS: &[(&str, &str)] = &[
    ("Error", "Exception"),
    ("DuplicateSectionError", "Error"),
    ("NoSectionError", "Error"),
    ("NoOptionError", "Error"),
];

pub(super) fn exception(name: &str, parent: &str) -> Statement {
    class_extending(name, &[parent], vec![])
}

pub(super) fn config_parser() -> Statement {
    config_parser_class("ConfigParser", false)
}

pub(super) fn raw_config_parser() -> Statement {
    config_parser_class("RawConfigParser", true)
}

pub(super) fn basic_interpolation() -> Statement {
    class("BasicInterpolation", vec![init(vec![], vec![])])
}

pub(super) fn extended_interpolation() -> Statement {
    class("ExtendedInterpolation", vec![init(vec![], vec![])])
}

fn op(op: BinOp, left: Expr, right: Expr) -> Expr {
    binary(op, left, right)
}

fn and(left: Expr, right: Expr) -> Expr {
    op(BinOp::And, left, right)
}

fn or_expr(left: Expr, right: Expr) -> Expr {
    op(BinOp::Or, left, right)
}

fn not_eq(left: Expr, right: Expr) -> Expr {
    op(BinOp::NotEq, left, right)
}

fn fp_append(text: Expr) -> Statement {
    expr_stmt(call(
        call_global("__py_attr_read", vec![ident("fp"), str_lit("write")]),
        vec![text],
    ))
}

fn first_char(expr: Expr) -> Expr {
    index(expr, Expr::int(0))
}

fn method_call(object: Expr, name: &str, args: Vec<Expr>) -> Expr {
    call(member(object, name), args)
}

fn config_parser_class(name: &str, raw: bool) -> Statement {
    class(
        name,
        vec![
            init(
                vec![
                    param("defaults", Some(null())),
                    param("dict_type", Some(null())),
                    param("allow_no_value", Some(bool_lit(false))),
                ],
                vec![
                    set_this("_sections", empty_dict()),
                    set_this("_defaults", empty_dict()),
                    set_this("_allow_no_value", ident("allow_no_value")),
                    set_this("_raw", bool_lit(raw)),
                    set_this("_optionxform", null()),
                    if_stmt(
                        is_not_none(ident("defaults")),
                        vec![for_in(
                            "key",
                            ident("defaults"),
                            vec![assign(
                                index(defaults_map(), ident("key")),
                                call_global("str", vec![index(ident("defaults"), ident("key"))]),
                            )],
                        )],
                    ),
                ],
            ),
            method(
                "read_string",
                vec![
                    param("text", Some(str_lit(""))),
                    param("source", Some(str_lit("<string>"))),
                ],
                vec![
                    set_this(
                        "_sections",
                        call_global(
                            "__py_config_parse",
                            vec![ident("text"), is_none(optionxform())],
                        ),
                    ),
                    expr_stmt(call(
                        member(ident("self"), "_merge_multiline_values"),
                        vec![ident("text")],
                    )),
                    if_stmt(
                        is_true(this_field("_allow_no_value")),
                        vec![expr_stmt(call(
                            member(ident("self"), "_merge_allow_no_value"),
                            vec![ident("text")],
                        ))],
                    ),
                ],
            ),
            property(
                "optionxform",
                vec![ret(optionxform())],
                param("value", Some(null())),
                vec![set_this("_optionxform", ident("value"))],
            ),
            method(
                "_normalize_option",
                vec![param("key", Some(null()))],
                vec![
                    if_stmt(
                        is_not_none(optionxform()),
                        vec![ret(call(optionxform(), vec![ident("key")]))],
                    ),
                    if_stmt(is_none(optionxform()), vec![ret(method_call(ident("key"), "lower", vec![]))]),
                    ret(ident("key")),
                ],
            ),
            method(
                "_merge_allow_no_value",
                vec![param("text", Some(str_lit("")))],
                vec![
                    assign(ident("__cfg_cur"), null()),
                    for_in(
                        "__cfg_raw",
                        method_call(ident("text"), "split", vec![str_lit("\n")]),
                        vec![
                            assign(ident("__cfg_line"), method_call(ident("__cfg_raw"), "strip", vec![])),
                            if_stmt(
                                and(
                                    and(
                                        not_eq(ident("__cfg_line"), str_lit("")),
                                        not_eq(first_char(ident("__cfg_line")), str_lit("#")),
                                    ),
                                    not_eq(first_char(ident("__cfg_line")), str_lit(";")),
                                ),
                                vec![
                                    if_stmt(
                                        op(BinOp::Eq, first_char(ident("__cfg_line")), str_lit("[")),
                                        vec![
                                            assign(
                                                ident("__cfg_end"),
                                                method_call(
                                                    ident("__cfg_line"),
                                                    "find",
                                                    vec![str_lit("]")],
                                                ),
                                            ),
                                            if_stmt(
                                                op(BinOp::LtEq, Expr::int(0), ident("__cfg_end")),
                                                vec![
                                                    assign(
                                                        ident("__cfg_cur"),
                                                        slice_range(
                                                            ident("__cfg_line"),
                                                            Expr::int(1),
                                                            ident("__cfg_end"),
                                                        ),
                                                    ),
                                                    if_stmt(
                                                        unary_not(contains(
                                                            sections(),
                                                            ident("__cfg_cur"),
                                                        )),
                                                        vec![assign(
                                                            sec_of(ident("__cfg_cur")),
                                                            empty_dict(),
                                                        )],
                                                    ),
                                                ],
                                            ),
                                        ],
                                    ),
                                    if_stmt(
                                        and(
                                            is_not_none(ident("__cfg_cur")),
                                            and(
                                                not_eq(first_char(ident("__cfg_line")), str_lit("[")),
                                                and(
                                                    op(
                                                        BinOp::Lt,
                                                        method_call(
                                                            ident("__cfg_line"),
                                                            "find",
                                                            vec![str_lit("=")],
                                                        ),
                                                        Expr::int(0),
                                                    ),
                                                    op(
                                                        BinOp::Lt,
                                                        method_call(
                                                            ident("__cfg_line"),
                                                            "find",
                                                            vec![str_lit(":")],
                                                        ),
                                                        Expr::int(0),
                                                    ),
                                                ),
                                            ),
                                        ),
                                        vec![assign(
                                            index(
                                                sec_of(ident("__cfg_cur")),
                                                call(
                                                    member(ident("self"), "_normalize_option"),
                                                    vec![ident("__cfg_line")],
                                                ),
                                            ),
                                            null(),
                                        )],
                                    ),
                                ],
                            ),
                        ],
                    ),
                    ret(null()),
                ],
            ),
            method(
                "_merge_multiline_values",
                vec![param("text", Some(str_lit("")))],
                vec![
                    assign(ident("__cfg_cur"), null()),
                    assign(ident("__cfg_last"), null()),
                    for_in(
                        "__cfg_raw",
                        method_call(ident("text"), "split", vec![str_lit("\n")]),
                        vec![
                            assign(
                                ident("__cfg_line"),
                                method_call(ident("__cfg_raw"), "strip", vec![]),
                            ),
                            if_stmt(
                                and(
                                    and(
                                        not_eq(ident("__cfg_line"), str_lit("")),
                                        not_eq(first_char(ident("__cfg_line")), str_lit("#")),
                                    ),
                                    not_eq(first_char(ident("__cfg_line")), str_lit(";")),
                                ),
                                vec![
                                    if_stmt(
                                        op(
                                            BinOp::Eq,
                                            first_char(ident("__cfg_line")),
                                            str_lit("["),
                                        ),
                                        vec![
                                            assign(
                                                ident("__cfg_end"),
                                                method_call(
                                                    ident("__cfg_line"),
                                                    "find",
                                                    vec![str_lit("]")],
                                                ),
                                            ),
                                            if_stmt(
                                                op(BinOp::LtEq, Expr::int(0), ident("__cfg_end")),
                                                vec![
                                                    assign(
                                                        ident("__cfg_cur"),
                                                        slice_range(
                                                            ident("__cfg_line"),
                                                            Expr::int(1),
                                                            ident("__cfg_end"),
                                                        ),
                                                    ),
                                                    assign(ident("__cfg_last"), null()),
                                                ],
                                            ),
                                        ],
                                    ),
                                    if_stmt(
                                        and(
                                            is_not_none(ident("__cfg_cur")),
                                            not_eq(first_char(ident("__cfg_line")), str_lit("[")),
                                        ),
                                        vec![
                                            if_stmt(
                                                and(
                                                    is_not_none(ident("__cfg_last")),
                                                    or_expr(
                                                        op(
                                                            BinOp::Eq,
                                                            method_call(
                                                                ident("__cfg_raw"),
                                                                "find",
                                                                vec![str_lit(" ")],
                                                            ),
                                                            Expr::int(0),
                                                        ),
                                                        op(
                                                            BinOp::Eq,
                                                            method_call(
                                                                ident("__cfg_raw"),
                                                                "find",
                                                                vec![str_lit("\t")],
                                                            ),
                                                            Expr::int(0),
                                                        ),
                                                    ),
                                                ),
                                                vec![assign(
                                                    index(
                                                        sec_of(ident("__cfg_cur")),
                                                        ident("__cfg_last"),
                                                    ),
                                                    add(
                                                        add(
                                                            index(
                                                                sec_of(ident("__cfg_cur")),
                                                                ident("__cfg_last"),
                                                            ),
                                                            str_lit("\n"),
                                                        ),
                                                        ident("__cfg_line"),
                                                    ),
                                                )],
                                            ),
                                            if_stmt(
                                                and(
                                                    op(
                                                        BinOp::LtEq,
                                                        Expr::int(0),
                                                        method_call(
                                                            ident("__cfg_line"),
                                                            "find",
                                                            vec![str_lit("=")],
                                                        ),
                                                    ),
                                                    unary_not(or_expr(
                                                        op(
                                                            BinOp::Eq,
                                                            method_call(
                                                                ident("__cfg_raw"),
                                                                "find",
                                                                vec![str_lit(" ")],
                                                            ),
                                                            Expr::int(0),
                                                        ),
                                                        op(
                                                            BinOp::Eq,
                                                            method_call(
                                                                ident("__cfg_raw"),
                                                                "find",
                                                                vec![str_lit("\t")],
                                                            ),
                                                            Expr::int(0),
                                                        ),
                                                    )),
                                                ),
                                                vec![
                                                    assign(
                                                        ident("__cfg_at"),
                                                        method_call(
                                                            ident("__cfg_line"),
                                                            "find",
                                                            vec![str_lit("=")],
                                                        ),
                                                    ),
                                                    assign(
                                                        ident("__cfg_last"),
                                                        call(
                                                            member(ident("self"), "_normalize_option"),
                                                            vec![method_call(
                                                                slice_range(
                                                                    ident("__cfg_line"),
                                                                    Expr::int(0),
                                                                    ident("__cfg_at"),
                                                                ),
                                                                "strip",
                                                                vec![],
                                                            )],
                                                        ),
                                                    ),
                                                ],
                                            ),
                                        ],
                                    ),
                                ],
                            ),
                        ],
                    ),
                    ret(null()),
                ],
            ),
            method(
                "read_dict",
                vec![param("d", Some(null()))],
                vec![for_in(
                    "name",
                    ident("d"),
                    vec![
                        assign(ident("target"), empty_dict()),
                        assign(ident("src"), index(ident("d"), ident("name"))),
                        for_in(
                            "key",
                            ident("src"),
                            vec![assign(
                                index(ident("target"), ident("key")),
                                call_global("str", vec![index(ident("src"), ident("key"))]),
                            )],
                        ),
                        assign(sec_of(ident("name")), ident("target")),
                    ],
                )],
            ),
            method(
                "sections",
                vec![],
                vec![
                    assign(ident("result"), list_of(vec![])),
                    for_in(
                        "k",
                        sections(),
                        vec![expr_stmt(call(
                            member(ident("result"), "append"),
                            vec![ident("k")],
                        ))],
                    ),
                    ret(ident("result")),
                ],
            ),
            method(
                "has_section",
                vec![param("sec", Some(null()))],
                vec![ret(contains(sections(), ident("sec")))],
            ),
            method(
                "add_section",
                vec![param("sec", Some(null()))],
                vec![
                    if_stmt(
                        contains(sections(), ident("sec")),
                        vec![raise_new("DuplicateSectionError", vec![])],
                    ),
                    assign(sec_of(ident("sec")), empty_dict()),
                    ret(null()),
                ],
            ),
            method(
                "set",
                vec![
                    param("sec", Some(null())),
                    param("opt", Some(null())),
                    param("value", Some(null())),
                ],
                vec![
                    if_stmt(
                        unary_not(contains(sections(), ident("sec"))),
                        vec![raise_new("NoSectionError", vec![])],
                    ),
                    assign(
                        index(sec_of(ident("sec")), ident("opt")),
                        call_global("str", vec![ident("value")]),
                    ),
                    ret(null()),
                ],
            ),
            method(
                "has_option",
                vec![param("sec", Some(null())), param("opt", Some(null()))],
                vec![
                    if_stmt(
                        unary_not(contains(sections(), ident("sec"))),
                        vec![ret(contains(defaults_map(), ident("opt")))],
                    ),
                    if_stmt(
                        contains(defaults_map(), ident("opt")),
                        vec![ret(bool_lit(true))],
                    ),
                    ret(contains(sec_of(ident("sec")), ident("opt"))),
                ],
            ),
            method(
                "remove_option",
                vec![param("sec", Some(null())), param("opt", Some(null()))],
                vec![
                    if_stmt(
                        unary_not(contains(sections(), ident("sec"))),
                        vec![ret(bool_lit(false))],
                    ),
                    if_stmt(
                        contains(sec_of(ident("sec")), ident("opt")),
                        vec![
                            expr_stmt(method_call(
                                sec_of(ident("sec")),
                                "pop",
                                vec![ident("opt")],
                            )),
                            ret(bool_lit(true)),
                        ],
                    ),
                    ret(bool_lit(false)),
                ],
            ),
            method(
                "remove_section",
                vec![param("sec", Some(null()))],
                vec![
                    if_stmt(
                        contains(sections(), ident("sec")),
                        vec![
                            expr_stmt(method_call(sections(), "pop", vec![ident("sec")])),
                            ret(bool_lit(true)),
                        ],
                    ),
                    ret(bool_lit(false)),
                ],
            ),
            method(
                "options",
                vec![param("sec", Some(null()))],
                vec![
                    assign(ident("result"), list_of(vec![])),
                    for_in(
                        "k",
                        defaults_map(),
                        vec![expr_stmt(call(
                            member(ident("result"), "append"),
                            vec![ident("k")],
                        ))],
                    ),
                    for_in(
                        "k",
                        sec_of(ident("sec")),
                        vec![expr_stmt(call(
                            member(ident("result"), "append"),
                            vec![ident("k")],
                        ))],
                    ),
                    ret(ident("result")),
                ],
            ),
            method(
                "items",
                vec![param("sec", Some(null()))],
                vec![
                    assign(ident("result"), list_of(vec![])),
                    for_in(
                        "key",
                        defaults_map(),
                        vec![expr_stmt(call(
                            member(ident("result"), "append"),
                            vec![tuple_of(vec![
                                ident("key"),
                                index(defaults_map(), ident("key")),
                            ])],
                        ))],
                    ),
                    if_stmt(
                        contains(sections(), ident("sec")),
                        vec![for_in(
                            "key",
                            sec_of(ident("sec")),
                            vec![expr_stmt(call(
                                member(ident("result"), "append"),
                                vec![tuple_of(vec![
                                    ident("key"),
                                    index(sec_of(ident("sec")), ident("key")),
                                ])],
                            ))],
                        )],
                    ),
                    ret(ident("result")),
                ],
            ),
            method(
                "get",
                vec![
                    param("sec", Some(null())),
                    param("opt", Some(null())),
                    param("fallback", Some(str_lit(MISSING))),
                ],
                vec![
                    if_stmt(
                        contains(sections(), ident("sec")),
                        vec![if_stmt(
                            contains(sec_of(ident("sec")), ident("opt")),
                            vec![ret(call(
                                member(ident("self"), "_interpolate"),
                                vec![
                                    ident("sec"),
                                    index(sec_of(ident("sec")), ident("opt")),
                                ],
                            ))],
                        )],
                    ),
                    if_stmt(
                        contains(sections(), str_lit("DEFAULT")),
                        vec![if_stmt(
                            contains(sec_of(str_lit("DEFAULT")), ident("opt")),
                            vec![ret(call(
                                member(ident("self"), "_interpolate"),
                                vec![
                                    ident("sec"),
                                    index(sec_of(str_lit("DEFAULT")), ident("opt")),
                                ],
                            ))],
                        )],
                    ),
                    if_stmt(
                        contains(defaults_map(), ident("opt")),
                        vec![ret(call(
                            member(ident("self"), "_interpolate"),
                            vec![ident("sec"), index(defaults_map(), ident("opt"))],
                        ))],
                    ),
                    if_stmt(
                        binary(BinOp::Eq, ident("fallback"), str_lit(MISSING)),
                        vec![if_stmt(
                            unary_not(contains(sections(), ident("sec"))),
                            vec![raise_new("NoSectionError", vec![])],
                        )],
                    ),
                    if_stmt(
                        binary(BinOp::Eq, ident("fallback"), str_lit(MISSING)),
                        vec![raise_new("NoOptionError", vec![])],
                    ),
                    ret(ident("fallback")),
                ],
            ),
            method(
                "_interpolate",
                vec![param("sec", Some(null())), param("value", Some(null()))],
                vec![
                    if_stmt(is_true(this_field("_raw")), vec![ret(ident("value"))]),
                    if_stmt(is_none(ident("value")), vec![ret(ident("value"))]),
                    assign(ident("result"), call_global("str", vec![ident("value")])),
                    for_in(
                        "key",
                        defaults_map(),
                        vec![assign(
                            ident("result"),
                            call(
                                member(ident("result"), "replace"),
                                vec![
                                    add(add(str_lit("%("), ident("key")), str_lit(")s")),
                                    index(defaults_map(), ident("key")),
                                ],
                            ),
                        )],
                    ),
                    if_stmt(
                        contains(sections(), ident("sec")),
                        vec![for_in(
                            "key",
                            sec_of(ident("sec")),
                            vec![assign(
                                ident("result"),
                                call(
                                    member(ident("result"), "replace"),
                                    vec![
                                        add(add(str_lit("%("), ident("key")), str_lit(")s")),
                                        index(sec_of(ident("sec")), ident("key")),
                                    ],
                                ),
                            )],
                        )],
                    ),
                    for_in(
                        "section",
                        sections(),
                        vec![for_in(
                            "key",
                            sec_of(ident("section")),
                            vec![assign(
                                ident("result"),
                                call(
                                    member(ident("result"), "replace"),
                                    vec![
                                        add(
                                            add(
                                                add(str_lit("${"), ident("section")),
                                                str_lit(":"),
                                            ),
                                            add(ident("key"), str_lit("}")),
                                        ),
                                        index(sec_of(ident("section")), ident("key")),
                                    ],
                                ),
                            )],
                        )],
                    ),
                    ret(ident("result")),
                ],
            ),
            method(
                "getint",
                vec![
                    param("sec", Some(null())),
                    param("opt", Some(null())),
                    param("fallback", Some(null())),
                ],
                vec![ret(call_global(
                    "int",
                    vec![call(
                        member(ident("self"), "get"),
                        vec![ident("sec"), ident("opt"), ident("fallback")],
                    )],
                ))],
            ),
            method(
                "getfloat",
                vec![
                    param("sec", Some(null())),
                    param("opt", Some(null())),
                    param("fallback", Some(null())),
                ],
                vec![ret(call_global(
                    "float",
                    vec![call(
                        member(ident("self"), "get"),
                        vec![ident("sec"), ident("opt"), ident("fallback")],
                    )],
                ))],
            ),
            method(
                "getboolean",
                vec![
                    param("sec", Some(null())),
                    param("opt", Some(null())),
                    param("fallback", Some(null())),
                ],
                vec![
                    assign(
                        ident("v"),
                        call_global(
                            "str",
                            vec![call(
                                member(ident("self"), "get"),
                                vec![ident("sec"), ident("opt"), ident("fallback")],
                            )],
                        ),
                    ),
                    assign(ident("v"), call(member(ident("v"), "lower"), vec![])),
                    ret(binary(
                        BinOp::Or,
                        binary(
                            BinOp::Or,
                            binary(BinOp::Eq, ident("v"), str_lit("true")),
                            binary(BinOp::Eq, ident("v"), str_lit("1")),
                        ),
                        binary(
                            BinOp::Or,
                            binary(BinOp::Eq, ident("v"), str_lit("yes")),
                            binary(BinOp::Eq, ident("v"), str_lit("on")),
                        ),
                    )),
                ],
            ),
            method("defaults", vec![], vec![ret(defaults_map())]),
            method(
                "__getitem__",
                vec![param("sec", Some(null()))],
                vec![
                    if_stmt(
                        op(BinOp::Eq, ident("sec"), str_lit("DEFAULT")),
                        vec![
                            if_stmt(
                                unary_not(contains(sections(), str_lit("DEFAULT"))),
                                vec![assign(sec_of(str_lit("DEFAULT")), empty_dict())],
                            ),
                            ret(sec_of(str_lit("DEFAULT"))),
                        ],
                    ),
                    assign(ident("result"), empty_dict()),
                    for_in(
                        "key",
                        defaults_map(),
                        vec![assign(
                            index(ident("result"), ident("key")),
                            index(defaults_map(), ident("key")),
                        )],
                    ),
                    if_stmt(
                        contains(sections(), ident("sec")),
                        vec![for_in(
                            "key",
                            sec_of(ident("sec")),
                            vec![assign(
                                index(ident("result"), ident("key")),
                                index(sec_of(ident("sec")), ident("key")),
                            )],
                        )],
                    ),
                    ret(ident("result")),
                ],
            ),
            method(
                "__setitem__",
                vec![param("sec", Some(null())), param("values", Some(null()))],
                vec![
                    assign(ident("target"), empty_dict()),
                    for_in(
                        "key",
                        ident("values"),
                        vec![assign(
                            index(ident("target"), ident("key")),
                            call_global("str", vec![index(ident("values"), ident("key"))]),
                        )],
                    ),
                    assign(sec_of(ident("sec")), ident("target")),
                ],
            ),
            method(
                "write",
                vec![param("fp", Some(null()))],
                vec![for_in(
                    "sec",
                    sections(),
                    vec![
                        fp_append(binary(
                            BinOp::Add,
                            binary(BinOp::Add, str_lit("["), ident("sec")),
                            str_lit("]\n"),
                        )),
                        for_in(
                            "key",
                            sec_of(ident("sec")),
                            vec![fp_append(binary(
                                BinOp::Add,
                                binary(
                                    BinOp::Add,
                                    binary(BinOp::Add, ident("key"), str_lit(" = ")),
                                    index(sec_of(ident("sec")), ident("key")),
                                ),
                                str_lit("\n"),
                            ))],
                        ),
                    ],
                )],
            ),
            method(
                "__contains__",
                vec![param("sec", Some(null()))],
                vec![ret(contains(sections(), ident("sec")))],
            ),
        ],
    )
}

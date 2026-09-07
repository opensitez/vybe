//! `gettext` — null translations and catalog lookup.
//!
//! This is the CPython-compatible no-catalog baseline: null translations return
//! the original message, `GNUTranslations` honors a manually installed
//! `_catalog`/`plural`, and module helpers route through the same classes.

use super::builders::*;
use vybe_ast::{BinOp, ExprKind, Expression, LambdaBody, Param, Statement};

type Expr = Expression;

fn contains(haystack: Expr, needle: Expr) -> Expr {
    call_global("__py_contains__", vec![haystack, needle])
}

fn not_none(expr: Expr) -> Expr {
    is_not_none(expr)
}

fn fallback_call(method_name: &str, args: Vec<Expr>) -> Statement {
    ret(call(member(this_field("_fallback"), method_name), args))
}

fn plural_lambda() -> Expr {
    Expression::with_span(
        ExprKind::Lambda {
            params: vec![Param {
                name: "n".into(),
                type_hint: None,
                default: None,
                pass_by: vybe_ast::PassBy::Value,
                is_rest: false,
                is_kwargs: false,
                is_optional: false,
                is_nullable: false,
            }],
            body: LambdaBody::Expr(Box::new(call_global(
                "int",
                vec![binary(BinOp::NotEq, ident("n"), num(1.0))],
            ))),
            is_async: false,
            captures: vec![],
        },
        span(),
    )
}

fn null_init_body() -> Vec<Statement> {
    vec![
        set_this("_info", dict_str(vec![])),
        set_this("_charset", null()),
        set_this("_fallback", null()),
    ]
}

fn install_body() -> Vec<Statement> {
    vec![ret(null())]
}

pub(super) fn null_translations() -> Statement {
    class(
        "NullTranslations",
        vec![
            init(vec![param("fp", Some(null()))], null_init_body()),
            method(
                "add_fallback",
                vec![param("fallback", None)],
                vec![set_this("_fallback", ident("fallback")), ret(null())],
            ),
            method(
                "gettext",
                vec![param("message", None)],
                vec![
                    if_stmt(
                        not_none(this_field("_fallback")),
                        vec![fallback_call("gettext", vec![ident("message")])],
                    ),
                    ret(ident("message")),
                ],
            ),
            method(
                "ngettext",
                vec![
                    param("singular", None),
                    param("plural", None),
                    param("n", None),
                ],
                vec![
                    if_stmt(
                        not_none(this_field("_fallback")),
                        vec![fallback_call(
                            "ngettext",
                            vec![ident("singular"), ident("plural"), ident("n")],
                        )],
                    ),
                    ret(ternary(
                        binary(BinOp::Eq, ident("n"), num(1.0)),
                        ident("singular"),
                        ident("plural"),
                    )),
                ],
            ),
            method(
                "pgettext",
                vec![param("context", None), param("message", None)],
                vec![ret(call(
                    member(ident("self"), "gettext"),
                    vec![ident("message")],
                ))],
            ),
            method(
                "npgettext",
                vec![
                    param("context", None),
                    param("singular", None),
                    param("plural", None),
                    param("n", None),
                ],
                vec![ret(call(
                    member(ident("self"), "ngettext"),
                    vec![ident("singular"), ident("plural"), ident("n")],
                ))],
            ),
            method("info", vec![], vec![ret(this_field("_info"))]),
            method("charset", vec![], vec![ret(this_field("_charset"))]),
            method(
                "install",
                vec![param("names", Some(null()))],
                install_body(),
            ),
        ],
    )
}

fn catalog_lookup_body() -> Vec<Statement> {
    vec![
        if_stmt(
            contains(this_field("_catalog"), ident("message")),
            vec![ret(index(this_field("_catalog"), ident("message")))],
        ),
        if_stmt(
            not_none(this_field("_fallback")),
            vec![fallback_call("gettext", vec![ident("message")])],
        ),
        ret(ident("message")),
    ]
}

fn catalog_plural_body() -> Vec<Statement> {
    vec![
        assign(ident("__plural"), this_field("plural")),
        assign(
            ident("__idx"),
            call(ident("__plural"), vec![ident("n")]),
        ),
        assign(
            ident("__key"),
            add(
                add(str_lit("__py_tuple:"), ident("singular")),
                add(str_lit("\x1f"), call_global("str", vec![ident("__idx")])),
            ),
        ),
        if_stmt(
            contains(this_field("_catalog"), ident("__key")),
            vec![ret(index(this_field("_catalog"), ident("__key")))],
        ),
        if_stmt(
            not_none(this_field("_fallback")),
            vec![fallback_call(
                "ngettext",
                vec![ident("singular"), ident("plural"), ident("n")],
            )],
        ),
        ret(ternary(
            binary(BinOp::Eq, ident("n"), num(1.0)),
            ident("singular"),
            ident("plural"),
        )),
    ]
}

pub(super) fn gnu_translations() -> Statement {
    class_extending(
        "GNUTranslations",
        &["NullTranslations"],
        vec![
            init(
                vec![param("fp", Some(null()))],
                vec![
                    set_this("_info", dict_str(vec![])),
                    set_this("_charset", null()),
                    set_this("_fallback", null()),
                    set_this("_catalog", dict_str(vec![])),
                    set_this("plural", plural_lambda()),
                ],
            ),
            method(
                "gettext",
                vec![param("message", None)],
                catalog_lookup_body(),
            ),
            method(
                "ngettext",
                vec![
                    param("singular", None),
                    param("plural", None),
                    param("n", None),
                ],
                catalog_plural_body(),
            ),
        ],
    )
}

fn identity(name: &str, params: Vec<vybe_ast::Param>, value: Expr) -> Statement {
    function(name, params, vec![ret(value)])
}

fn raise_oserror() -> Statement {
    expr_stmt(call_global("__py_raise_OSError", vec![]))
}

pub(super) fn module_functions() -> Vec<Statement> {
    vec![
        identity("_", vec![param("message", None)], ident("message")),
        identity(
            "__py_gettext_gettext",
            vec![param("message", None)],
            ident("message"),
        ),
        identity(
            "__py_gettext_dgettext",
            vec![param("domain", None), param("message", None)],
            ident("message"),
        ),
        identity(
            "__py_gettext_pgettext",
            vec![param("context", None), param("message", None)],
            ident("message"),
        ),
        identity(
            "__py_gettext_dpgettext",
            vec![
                param("domain", None),
                param("context", None),
                param("message", None),
            ],
            ident("message"),
        ),
        function(
            "__py_gettext_ngettext",
            vec![
                param("singular", None),
                param("plural", None),
                param("n", None),
            ],
            vec![ret(ternary(
                binary(BinOp::Eq, ident("n"), num(1.0)),
                ident("singular"),
                ident("plural"),
            ))],
        ),
        function(
            "__py_gettext_dngettext",
            vec![
                param("domain", None),
                param("singular", None),
                param("plural", None),
                param("n", None),
            ],
            vec![ret(call_global(
                "__py_gettext_ngettext",
                vec![ident("singular"), ident("plural"), ident("n")],
            ))],
        ),
        function(
            "__py_gettext_npgettext",
            vec![
                param("context", None),
                param("singular", None),
                param("plural", None),
                param("n", None),
            ],
            vec![ret(call_global(
                "__py_gettext_ngettext",
                vec![ident("singular"), ident("plural"), ident("n")],
            ))],
        ),
        function(
            "__py_gettext_dnpgettext",
            vec![
                param("domain", None),
                param("context", None),
                param("singular", None),
                param("plural", None),
                param("n", None),
            ],
            vec![ret(call_global(
                "__py_gettext_ngettext",
                vec![ident("singular"), ident("plural"), ident("n")],
            ))],
        ),
        identity(
            "__py_gettext_bindtextdomain",
            vec![param("domain", None), param("localedir", Some(null()))],
            ident("localedir"),
        ),
        identity(
            "__py_gettext_textdomain",
            vec![param("domain", Some(str_lit("messages")))],
            ident("domain"),
        ),
        identity(
            "__py_gettext_find",
            vec![
                param("domain", None),
                param("localedir", Some(null())),
                param("languages", Some(null())),
                param("all", Some(bool_lit(false))),
            ],
            null(),
        ),
        function(
            "__py_gettext_translation",
            vec![
                param("domain", None),
                param("localedir", Some(null())),
                param("languages", Some(null())),
                param("class_", Some(null())),
                param("fallback", Some(bool_lit(false))),
                param("codeset", Some(null())),
            ],
            vec![
                if_stmt(
                    binary(BinOp::Eq, ident("fallback"), bool_lit(true)),
                    vec![ret(new("NullTranslations", vec![]))],
                ),
                raise_oserror(),
            ],
        ),
        function(
            "__py_gettext_install",
            vec![
                param("domain", None),
                param("localedir", Some(null())),
                param("codeset", Some(null())),
                param("names", Some(null())),
            ],
            vec![ret(null())],
        ),
    ]
}

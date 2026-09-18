//! `warnings` — the warning category hierarchy plus `catch_warnings`.
//!
//! Eleven of the thirteen classes are pure INHERITANCE (`class UserWarning(
//! Warning): pass`), which is the cheapest possible thing to declare and the
//! thing a prelude bought least by parsing. The parent chain is not decoration:
//! it is what puts the ancestor into the `__types` chain `compile_class`
//! stamps, and that chain is half of what `reflection::emit_is_instance_of`
//! unions with the rtt to answer `except UserWarning` — so
//! `except Warning` catching a `DeprecationWarning` is exactly this declaration.

use super::builders::*;
use vybe_ast::{BinOp, Expression, Statement, StmtKind};

/// `Warning` and its ten standard subclasses, in declaration order — a class
/// must follow the one it extends so the ancestor's MRO is resolved when the
/// child's chain is stamped.
pub(super) const CATEGORIES: &[(&str, &str)] = &[
    ("Warning", "Exception"),
    ("UserWarning", "Warning"),
    ("DeprecationWarning", "Warning"),
    ("PendingDeprecationWarning", "Warning"),
    ("SyntaxWarning", "Warning"),
    ("RuntimeWarning", "Warning"),
    ("FutureWarning", "Warning"),
    ("ImportWarning", "Warning"),
    ("UnicodeWarning", "Warning"),
    ("BytesWarning", "Warning"),
    ("ResourceWarning", "Warning"),
];

/// One category class. `CATEGORIES` is the single list; `mod.rs` turns each row
/// into a `CORE_CLASSES` entry, so the hierarchy is stated once.
pub(super) fn category(name: &'static str, parent: &'static str) -> Statement {
    class_extending(
        name,
        &[parent],
        vec![
            init(
                vec![param("message", Some(str_lit("")))],
                vec![set_this("message", ident("message"))],
            ),
            method("__str__", vec![], vec![ret(this_field("message"))]),
        ],
    )
}

/// One recorded warning — what `catch_warnings(record=True)` appends.
pub(super) fn warning_record() -> Statement {
    class(
        "__WarningRecord",
        vec![init(
            vec![
                param("message", None),
                param("category", None),
                param("filename", Some(str_lit(""))),
                param("lineno", Some(Expression::int(0))),
            ],
            vec![
                set_this(
                    "message",
                    warning_category_instance(ident("message"), ident("category")),
                ),
                set_this("category", ident("category")),
                set_this("filename", ident("filename")),
                set_this("lineno", ident("lineno")),
            ],
        )],
    )
}

/// The `catch_warnings` context manager. `__enter__` / `__exit__` are declared
/// as ordinary dunders: python's `protocol.rs` maps both onto their
/// `ProtocolSlot`, so `with warnings.catch_warnings(record=True) as w:` binds
/// through the shared machinery with nothing context-manager-specific here.
pub(super) fn catch_warnings() -> Statement {
    class(
        "__CatchWarnings",
        vec![
            init(
                vec![param("record", None)],
                vec![
                    set_this("record", ident("record")),
                    set_this("entries", call_global("list", vec![])),
                ],
            ),
            method(
                "__enter__",
                vec![],
                vec![
                    if_stmt(
                        this_field("record"),
                        vec![
                            set_this(
                                "filters",
                                call_global("list", vec![ident("__py_warnings_filters")]),
                            ),
                            assign(ident("__vybe_warn_log"), this_field("entries")),
                            ret(this_field("entries")),
                        ],
                    ),
                    set_this(
                        "filters",
                        call_global("list", vec![ident("__py_warnings_filters")]),
                    ),
                    assign(ident("__vybe_warn_log"), Expression::null()),
                    ret(Expression::null()),
                ],
            ),
            method(
                "__exit__",
                vec![param("a", None), param("b", None), param("c", None)],
                vec![
                    expr_stmt(call(
                        member(ident("__py_warnings_filters"), "clear"),
                        vec![],
                    )),
                    expr_stmt(call(
                        member(ident("__py_warnings_filters"), "extend"),
                        vec![this_field("filters")],
                    )),
                    assign(ident("__vybe_warn_log"), Expression::null()),
                    ret(bool_lit(false)),
                ],
            ),
        ],
    )
}

/// The module-level surface. `warn` records into the global the context manager
/// binds — a module-level global rather than state on a module OBJECT, because
/// the module object was the prelude's own invention and nothing else needs it.
pub(super) fn module_functions() -> Vec<Statement> {
    vec![
        assign_global("__vybe_warn_log", Expression::null()),
        assign_global("__py_warnings_filters", list_of(vec![])),
        assign_global("__py_warnings_onceregistry", dict_of(vec![])),
        function(
            "__py_warning_action",
            vec![param("category", None)],
            vec![
                for_in(
                    "__filter",
                    ident("__py_warnings_filters"),
                    vec![if_stmt(
                        binary(
                            BinOp::Or,
                            call_global(
                                "__py_is__",
                                vec![
                                    index(ident("__filter"), Expression::int(2)),
                                    ident("Warning"),
                                ],
                            ),
                            call_global(
                                "__py_is__",
                                vec![
                                    index(ident("__filter"), Expression::int(2)),
                                    ident("category"),
                                ],
                            ),
                        ),
                        vec![ret(index(ident("__filter"), Expression::int(0)))],
                    )],
                ),
                ret(str_lit("default")),
            ],
        ),
        function(
            "warn",
            vec![
                param("message", None),
                param("category", Some(Expression::null())),
                param("stacklevel", Some(Expression::int(1))),
                param("source", Some(Expression::null())),
            ],
            vec![
                assign(
                    ident("__cat"),
                    ternary(
                        is_none(ident("category")),
                        ident("UserWarning"),
                        ident("category"),
                    ),
                ),
                assign(
                    ident("__action"),
                    call_global("__py_warning_action", vec![ident("__cat")]),
                ),
                if_stmt(
                    binary(BinOp::Eq, ident("__action"), str_lit("ignore")),
                    vec![ret(Expression::null())],
                ),
                if_stmt(
                    binary(BinOp::Eq, ident("__action"), str_lit("error")),
                    vec![throw_expr(warning_category_instance(
                        ident("message"),
                        ident("__cat"),
                    ))],
                ),
                if_stmt(
                    is_not_none(ident("__vybe_warn_log")),
                    vec![expr_stmt(call(
                        member(ident("__vybe_warn_log"), "append"),
                        vec![new(
                            "__WarningRecord",
                            vec![
                                ident("message"),
                                ident("__cat"),
                                str_lit(""),
                                Expression::int(0),
                            ],
                        )],
                    ))],
                ),
            ],
        ),
        function(
            "warn_explicit",
            vec![
                param("message", None),
                param("category", Some(Expression::null())),
                param("filename", Some(str_lit(""))),
                param("lineno", Some(Expression::int(0))),
            ],
            vec![
                assign(
                    ident("__cat"),
                    ternary(
                        is_none(ident("category")),
                        ident("UserWarning"),
                        ident("category"),
                    ),
                ),
                assign(
                    ident("__action"),
                    call_global("__py_warning_action", vec![ident("__cat")]),
                ),
                if_stmt(
                    binary(BinOp::Eq, ident("__action"), str_lit("ignore")),
                    vec![ret(Expression::null())],
                ),
                if_stmt(
                    binary(BinOp::Eq, ident("__action"), str_lit("error")),
                    vec![throw_expr(warning_category_instance(
                        ident("message"),
                        ident("__cat"),
                    ))],
                ),
                if_stmt(
                    is_not_none(ident("__vybe_warn_log")),
                    vec![expr_stmt(call(
                        member(ident("__vybe_warn_log"), "append"),
                        vec![new(
                            "__WarningRecord",
                            vec![
                                ident("message"),
                                ident("__cat"),
                                ident("filename"),
                                ident("lineno"),
                            ],
                        )],
                    ))],
                ),
            ],
        ),
        function(
            "showwarning",
            vec![
                param("message", None),
                param("category", Some(Expression::null())),
                param("filename", Some(str_lit(""))),
                param("lineno", Some(Expression::int(0))),
                param("file", Some(Expression::null())),
                param("line", Some(Expression::null())),
            ],
            vec![if_stmt(
                is_not_none(ident("file")),
                vec![expr_stmt(call(
                    member(ident("file"), "write"),
                    vec![add(
                        call_global("str", vec![ident("message")]),
                        str_lit("\n"),
                    )],
                ))],
            )],
        ),
        function(
            "filterwarnings",
            vec![
                param("action", Some(str_lit("default"))),
                param("message", Some(str_lit(""))),
                param("category", Some(ident("Warning"))),
                param("module", Some(str_lit(""))),
                param("lineno", Some(Expression::int(0))),
                param("append", Some(bool_lit(false))),
            ],
            vec![expr_stmt(call(
                member(ident("__py_warnings_filters"), "insert"),
                vec![
                    Expression::int(0),
                    list_of(vec![
                        ident("action"),
                        ident("message"),
                        ident("category"),
                        ident("module"),
                        ident("lineno"),
                    ]),
                ],
            ))],
        ),
        function(
            "simplefilter",
            vec![
                param("action", Some(str_lit("default"))),
                param("category", Some(ident("Warning"))),
                param("lineno", Some(Expression::int(0))),
                param("append", Some(bool_lit(false))),
            ],
            vec![expr_stmt(call(
                member(ident("__py_warnings_filters"), "insert"),
                vec![
                    Expression::int(0),
                    list_of(vec![
                        ident("action"),
                        str_lit(""),
                        ident("category"),
                        str_lit(""),
                        ident("lineno"),
                    ]),
                ],
            ))],
        ),
        function(
            "resetwarnings",
            vec![],
            vec![expr_stmt(call(
                member(ident("__py_warnings_filters"), "clear"),
                vec![],
            ))],
        ),
        function("_filters_mutated", vec![], vec![]),
        function(
            "catch_warnings",
            vec![param("record", Some(bool_lit(false)))],
            vec![ret(new("__CatchWarnings", vec![ident("record")]))],
        ),
        function(
            "formatwarning",
            vec![param("a", Some(Expression::null()))],
            vec![ret(str_lit(""))],
        ),
    ]
}

/// A module-level `name = value` binding.
fn assign_global(name: &str, value: Expression) -> Statement {
    assign(ident(name), value)
}

fn throw_expr(expr: Expression) -> Statement {
    Statement::with_span(
        StmtKind::Throw {
            expr: Some(expr),
            cause: None,
        },
        span(),
    )
}

fn warning_category_instance(message: Expression, category: Expression) -> Expression {
    CATEGORIES.iter().rev().fold(
        new("UserWarning", vec![message.clone()]),
        |else_, (name, _)| {
            ternary(
                call_global("__py_is__", vec![category.clone(), ident(name)]),
                new(name, vec![message.clone()]),
                else_,
            )
        },
    )
}

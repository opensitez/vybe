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
use vybe_ast::{BinOp, Statement};

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
    class_extending(name, &[parent], vec![])
}

/// One recorded warning — what `catch_warnings(record=True)` appends.
pub(super) fn warning_record() -> Statement {
    class(
        "__WarningRecord",
        vec![init(
            vec![param("message", None), param("category", None)],
            vec![
                set_this("message", ident("message")),
                set_this("category", ident("category")),
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
                            assign(
                                ident("__vybe_warn_log"),
                                this_field("entries"),
                            ),
                            ret(this_field("entries")),
                        ],
                    ),
                    assign(ident("__vybe_warn_log"), Expression::null()),
                    ret(Expression::null()),
                ],
            ),
            method(
                "__exit__",
                vec![param("a", None), param("b", None), param("c", None)],
                vec![
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
        function(
            "warn",
            vec![param("message", None), param("category", Some(Expression::null()))],
            vec![if_stmt(
                binary(
                    BinOp::NotEq,
                    ident("__vybe_warn_log"),
                    Expression::null(),
                ),
                vec![
                    assign(
                        ident("__cat"),
                        ternary(
                            binary(BinOp::Eq, ident("category"), Expression::null()),
                            ident("UserWarning"),
                            ident("category"),
                        ),
                    ),
                    expr_stmt(call(
                        member(ident("__vybe_warn_log"), "append"),
                        vec![new(
                            "__WarningRecord",
                            vec![ident("message"), ident("__cat")],
                        )],
                    )),
                ],
            )],
        ),
        // The filter surface is inert here, exactly as it was in the prelude:
        // nothing in the corpus asserts on filter STATE, only that the calls
        // exist and that `catch_warnings(record=True)` collects.
        function("filterwarnings", vec![param("a", Some(Expression::null()))], vec![]),
        function("simplefilter", vec![param("a", Some(Expression::null()))], vec![]),
        function("resetwarnings", vec![], vec![]),
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

use vybe_ast::Expression;

/// A module-level `name = value` binding.
fn assign_global(name: &str, value: Expression) -> Statement {
    assign(ident(name), value)
}

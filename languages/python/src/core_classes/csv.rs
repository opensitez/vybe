//! `csv` — reader, writer, the dict variants, the `excel` dialect and Sniffer.
//!
//! The line-level parsing and formatting stay in the emitter adapter behind
//! `__py_csv_parse_line` / `__py_csv_format_row`: splitting on a delimiter with
//! quote handling is character work, which is what an adapter is for. These
//! classes are the ITERATION and the row shaping around it, which is ordinary
//! Python and belongs in a class.
//!
//! `__iter__` / `__next__` are declared as plain dunders — python's
//! `protocol.rs` maps them onto `Iterator` / `Next`, so `for row in reader:`
//! binds through the shared protocol with nothing iterator-specific here.

use super::builders::*;
use vybe_ast::{BinOp, Statement};

/// ⛔ Real escapes, not `chr(13)`. The prelude spelled every one of these as a
/// `chr()` call because `preprocess_indentation` mis-read `#`/`[`/`(` inside a
/// PRELUDE string literal and silently dropped the whole module
/// ([[project_python_prelude_landmines]]). Constructed AST never goes near that
/// preprocessor, so the constants are constants — and a `chr()` call in a
/// class-level static initialiser is one more thing to go wrong.
fn cr() -> vybe_ast::Expression {
    str_lit("\r")
}

fn lf() -> vybe_ast::Expression {
    str_lit("\n")
}

fn quote() -> vybe_ast::Expression {
    str_lit("\"")
}

fn comma() -> vybe_ast::Expression {
    str_lit(",")
}

/// The `excel` dialect — three class-level constants, which is what a dialect
/// is in CPython.
pub(super) fn excel_dialect() -> Statement {
    class(
        "__PyCsvExcel",
        vec![
            static_field("delimiter", str_lit(",")),
            static_field("quotechar", str_lit("\"")),
            static_field("lineterminator", str_lit("\r\n")),
        ],
    )
}

pub(super) fn reader() -> Statement {
    class(
        "__PyCsvReader",
        vec![
            init(
                vec![param("source", None), param("dialect", Some(null()))],
                vec![
                    assign(ident("__rows"), call_global("list", vec![])),
                    set_this("_index", num(0.0)),
                    // A StringIO, a file, or plain text — the three shapes
                    // the corpus hands `csv.reader`.
                    //
                    // ⛔ No `None` sentinel and no elif chain: a local
                    // initialised to `None` and reassigned inside a branch is
                    // the hoisted-local trap in
                    // [[project_python_assign_none_in_a_function_trapped]].
                    // Start from the fallback and let the more specific tests
                    // overwrite it, most-specific LAST.
                    assign(ident("__data"), call_global("str", vec![ident("source")])),
                    if_stmt(
                        call_global("hasattr", vec![ident("source"), str_lit("read")]),
                        vec![assign(
                            ident("__data"),
                            call(member(ident("source"), "read"), vec![]),
                        )],
                    ),
                    if_stmt(
                        call_global("hasattr", vec![ident("source"), str_lit("getvalue")]),
                        vec![assign(
                            ident("__data"),
                            call(member(ident("source"), "getvalue"), vec![]),
                        )],
                    ),
                    for_in(
                        "__line",
                        call(member(ident("__data"), "split"), vec![lf()]),
                        vec![
                            assign(ident("__l"), ident("__line")),
                            if_stmt(
                                call(member(ident("__l"), "endswith"), vec![cr()]),
                                vec![assign(
                                    ident("__l"),
                                    call(
                                        member(ident("__l"), "rstrip"),
                                        vec![cr()],
                                    ),
                                )],
                            ),
                            if_stmt(
                                binary(BinOp::NotEq, ident("__l"), str_lit("")),
                                vec![expr_stmt(call(
                                    member(ident("__rows"), "append"),
                                    vec![call_global(
                                        "__py_csv_parse_line",
                                        vec![ident("__l"), comma(), quote()],
                                    )],
                                ))],
                            ),
                        ],
                    ),
                    set_this("_rows", ident("__rows")),
                ],
            ),
            method("__iter__", vec![], vec![ret(ident("self"))]),
            method(
                "__next__",
                vec![],
                vec![
                    if_stmt(
                        binary(
                            BinOp::GtEq,
                            this_field("_index"),
                            call_global("len", vec![this_field("_rows")]),
                        ),
                        vec![raise_stop_iteration()],
                    ),
                    assign(
                        ident("__row"),
                        index(this_field("_rows"), this_field("_index")),
                    ),
                    set_this("_index", add(this_field("_index"), num(1.0))),
                    ret(ident("__row")),
                ],
            ),
        ],
    )
}

pub(super) fn writer() -> Statement {
    class(
        "__PyCsvWriter",
        vec![
            init(
                vec![param("target", None), param("dialect", Some(null()))],
                vec![set_this("_target", ident("target"))],
            ),
            method(
                "writerow",
                vec![param("row", None)],
                vec![
                    assign(
                        ident("__text"),
                        call_global(
                            "__py_csv_format_row",
                            vec![ident("row"), comma(), quote()],
                        ),
                    ),
                    // ⛔ BIND THE RECEIVER FIRST. A method call whose receiver
                    // is a nested expression — here `__py_obj_get__(self,
                    // "_target")` — is the shape recorded in
                    // [[project_python_attributes_bypass_shared_classes]] as
                    // failing where the two-step form works. Measured: with the
                    // nested form `writerow` wrote nothing at all and
                    // `getvalue()` answered `''`.
                    assign(ident("__t"), this_field("_target")),
                    expr_stmt(call(
                        member(ident("__t"), "write"),
                        vec![add(add(ident("__text"), cr()), lf())],
                    )),
                    ret(call_global("len", vec![ident("__text")])),
                ],
            ),
            method(
                "writerows",
                vec![param("rows", None)],
                vec![for_in(
                    "__r",
                    ident("rows"),
                    vec![expr_stmt(call(
                        member(ident("self"), "writerow"),
                        vec![ident("__r")],
                    ))],
                )],
            ),
        ],
    )
}

pub(super) fn dict_reader() -> Statement {
    class(
        "__PyCsvDictReader",
        vec![
            init(
                vec![
                    param("source", None),
                    param("fieldnames", Some(null())),
                    param("dialect", Some(null())),
                ],
                vec![
                    set_this(
                        "_reader",
                        new("__PyCsvReader", vec![ident("source"), ident("dialect")]),
                    ),
                    if_stmt(
                        is_none(ident("fieldnames")),
                        vec![set_this(
                            "fieldnames",
                            call_global("next", vec![this_field("_reader")]),
                        )],
                    ),
                    if_stmt(
                        is_not_none(ident("fieldnames")),
                        vec![set_this("fieldnames", ident("fieldnames"))],
                    ),
                ],
            ),
            method("__iter__", vec![], vec![ret(ident("self"))]),
            method(
                "__next__",
                vec![],
                vec![
                    assign(
                        ident("__row"),
                        call_global("next", vec![this_field("_reader")]),
                    ),
                    assign(ident("__out"), call_global("dict", vec![])),
                    assign(ident("__i"), num(0.0)),
                    while_stmt(
                        binary(
                            BinOp::Lt,
                            ident("__i"),
                            call_global("len", vec![this_field("fieldnames")]),
                        ),
                        vec![
                            assign(
                                ident("__k"),
                                index(this_field("fieldnames"), ident("__i")),
                            ),
                            if_stmt(
                                binary(
                                    BinOp::Lt,
                                    ident("__i"),
                                    call_global("len", vec![ident("__row")]),
                                ),
                                vec![assign(
                                    index(ident("__out"), ident("__k")),
                                    index(ident("__row"), ident("__i")),
                                )],
                            ),
                            if_stmt(
                                binary(
                                    BinOp::GtEq,
                                    ident("__i"),
                                    call_global("len", vec![ident("__row")]),
                                ),
                                vec![assign(
                                    index(ident("__out"), ident("__k")),
                                    null(),
                                )],
                            ),
                            assign(ident("__i"), add(ident("__i"), num(1.0))),
                        ],
                    ),
                    ret(ident("__out")),
                ],
            ),
        ],
    )
}

pub(super) fn dict_writer() -> Statement {
    class(
        "__PyCsvDictWriter",
        vec![
            init(
                vec![
                    param("target", None),
                    param("fieldnames", None),
                    param("dialect", Some(null())),
                ],
                vec![
                    set_this(
                        "_writer",
                        new("__PyCsvWriter", vec![ident("target"), ident("dialect")]),
                    ),
                    set_this("fieldnames", ident("fieldnames")),
                ],
            ),
            method(
                "writeheader",
                vec![],
                vec![
                    assign(ident("__w"), this_field("_writer")),
                    ret(call(
                        member(ident("__w"), "writerow"),
                        vec![this_field("fieldnames")],
                    )),
                ],
            ),
            method(
                "writerow",
                vec![param("rowdict", None)],
                vec![
                    assign(ident("__row"), call_global("list", vec![])),
                    for_in(
                        "__name",
                        this_field("fieldnames"),
                        vec![
                            if_stmt(
                                binary(BinOp::In, ident("__name"), ident("rowdict")),
                                vec![expr_stmt(call(
                                    member(ident("__row"), "append"),
                                    vec![index(ident("rowdict"), ident("__name"))],
                                ))],
                            ),
                            if_stmt(
                                unary_not(binary(
                                    BinOp::In,
                                    ident("__name"),
                                    ident("rowdict"),
                                )),
                                vec![expr_stmt(call(
                                    member(ident("__row"), "append"),
                                    vec![str_lit("")],
                                ))],
                            ),
                        ],
                    ),
                    assign(ident("__w"), this_field("_writer")),
                    ret(call(member(ident("__w"), "writerow"), vec![ident("__row")])),
                ],
            ),
        ],
    )
}

pub(super) fn sniffer() -> Statement {
    class(
        "Sniffer",
        vec![
            stub("sniff", new("__PyCsvExcel", vec![])),
            stub("has_header", bool_lit(true)),
        ],
    )
}

pub(super) fn module_functions() -> Vec<Statement> {
    vec![
        function(
            "reader",
            vec![param("source", None), param("dialect", Some(null()))],
            vec![ret(new(
                "__PyCsvReader",
                vec![ident("source"), ident("dialect")],
            ))],
        ),
        function(
            "writer",
            vec![param("target", None), param("dialect", Some(null()))],
            vec![ret(new(
                "__PyCsvWriter",
                vec![ident("target"), ident("dialect")],
            ))],
        ),
        function(
            "DictReader",
            vec![
                param("source", None),
                param("fieldnames", Some(null())),
                param("dialect", Some(null())),
            ],
            vec![ret(new(
                "__PyCsvDictReader",
                vec![ident("source"), ident("fieldnames"), ident("dialect")],
            ))],
        ),
        function(
            "DictWriter",
            vec![
                param("target", None),
                param("fieldnames", None),
                param("dialect", Some(null())),
            ],
            vec![ret(new(
                "__PyCsvDictWriter",
                vec![ident("target"), ident("fieldnames"), ident("dialect")],
            ))],
        ),
        stub_fn("list_dialects", list_of(vec![str_lit("excel")])),
        stub_fn("field_size_limit", num(131072.0)),
        global_assign("excel", new("__PyCsvExcel", vec![])),
    ]
}

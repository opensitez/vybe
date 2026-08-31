//! `pathlib` — `PurePath`, its two flavours, and `Path` on the real filesystem.
//!
//! ▶▶ The prelude's `Path` was STUBS: `read_text` answered `''`, `iterdir`
//! answered `[]`, `mkdir`/`unlink`/`rename` answered `None`. Every one of those
//! is a `common:filesystem.*` primitive that already exists and that python's
//! own `os` surface has been using all along — pathlib simply never reached
//! them. So this conversion is also the fix: `Path.read_text()` reads the file.
//!
//! The pure-path half is string work and stays string work. It delegates the
//! two genuinely fiddly pieces — normalisation and glob matching — to declared
//! module helpers rather than re-deriving them per method.

use super::builders::*;
use vybe_ast::{BinOp, Statement};

fn slash() -> vybe_ast::Expression {
    str_lit("/")
}

/// `self._s` bound to a local — the receiver rule.
fn s_local() -> Statement {
    assign(ident("__s"), this_field("_s"))
}

/// `len(x)`
fn len_of(e: vybe_ast::Expression) -> vybe_ast::Expression {
    call_global("len", vec![e])
}

/// The last non-empty segment of `rest`, as a loop — `name`/`parts` share it.
fn split_segments(src: vybe_ast::Expression) -> Statement {
    assign(
        ident("__seg"),
        call(member(src, "split"), vec![slash()]),
    )
}

pub(super) fn pure_path() -> Statement {
    class(
        "PurePath",
        vec![
            init(
                vec![param("p", Some(str_lit("")))],
                vec![
                    set_this(
                        "_s",
                        call_global(
                            "_pp_norm",
                            vec![call_global("_pp_str", vec![ident("p")])],
                        ),
                    ),
                    // ⛔ A FIELD, not a method. `self._is_win()` inside a
                    // PROPERTY accessor does not bind its receiver — the
                    // ambient-`this` branch runs before scope resolution, so an
                    // accessor chunk's receiver parameter is looked past
                    // ([[project_this_ambient_branch_beats_a_bound_receiver]]).
                    // `p.drive` threw on exactly that call. Reading a field is
                    // fine; CALLING a method on `self` is not.
                    set_this("_win", bool_lit(false)),
                ],
            ),
            method("_is_win", vec![], vec![ret(this_field("_win"))]),
            method(
                "_make",
                vec![param("s", None)],
                vec![ret(new("PurePath", vec![ident("s")]))],
            ),
            // ── the pure-string properties ──────────────────────────────
            getter(
                "drive",
                vec![
                    if_stmt(
                        unary_not(this_field("_win")),
                        vec![ret(str_lit(""))],
                    ),
                    s_local(),
                    if_stmt(
                        binary(
                            BinOp::And,
                            binary(BinOp::GtEq, len_of(ident("__s")), num(2.0)),
                            binary(
                                BinOp::Eq,
                                index(ident("__s"), num(1.0)),
                                str_lit(":"),
                            ),
                        ),
                        vec![ret(slice_range(ident("__s"), num(0.0), num(2.0)))],
                    ),
                    ret(str_lit("")),
                ],
            ),
            getter(
                "root",
                vec![
                    assign(ident("__d"), this_field("drive")),
                    assign(
                        ident("__rest"),
                        slice_from(this_field("_s"), len_of(ident("__d"))),
                    ),
                    if_stmt(
                        binary(
                            BinOp::And,
                            binary(BinOp::GtEq, len_of(ident("__rest")), num(1.0)),
                            binary(BinOp::Eq, index(ident("__rest"), num(0.0)), slash()),
                        ),
                        vec![ret(slash())],
                    ),
                    ret(str_lit("")),
                ],
            ),
            getter(
                "anchor",
                vec![ret(add(this_field("drive"), this_field("root")))],
            ),
            getter(
                "name",
                vec![
                    assign(
                        ident("__rest"),
                        slice_from(this_field("_s"), len_of(this_field("anchor"))),
                    ),
                    assign(ident("__last"), str_lit("")),
                    split_segments(ident("__rest")),
                    for_in(
                        "__c",
                        ident("__seg"),
                        vec![if_stmt(
                            binary(BinOp::NotEq, ident("__c"), str_lit("")),
                            vec![assign(ident("__last"), ident("__c"))],
                        )],
                    ),
                    ret(ident("__last")),
                ],
            ),
            getter(
                "stem",
                vec![
                    assign(ident("__nm"), this_field("name")),
                    assign(
                        ident("__i"),
                        call(member(ident("__nm"), "rfind"), vec![str_lit(".")]),
                    ),
                    if_stmt(
                        binary(BinOp::Gt, ident("__i"), num(0.0)),
                        vec![ret(slice_range(ident("__nm"), num(0.0), ident("__i")))],
                    ),
                    ret(ident("__nm")),
                ],
            ),
            getter(
                "suffix",
                vec![
                    assign(ident("__nm"), this_field("name")),
                    assign(
                        ident("__i"),
                        call(member(ident("__nm"), "rfind"), vec![str_lit(".")]),
                    ),
                    if_stmt(
                        binary(BinOp::Gt, ident("__i"), num(0.0)),
                        vec![ret(slice_from(ident("__nm"), ident("__i")))],
                    ),
                    ret(str_lit("")),
                ],
            ),
            getter(
                "suffixes",
                vec![
                    assign(ident("__nm"), this_field("name")),
                    assign(ident("__out"), call_global("list", vec![])),
                    if_stmt(
                        call(member(ident("__nm"), "endswith"), vec![str_lit(".")]),
                        vec![ret(ident("__out"))],
                    ),
                    assign(
                        ident("__pieces"),
                        call(member(ident("__nm"), "split"), vec![str_lit(".")]),
                    ),
                    assign(ident("__i"), num(1.0)),
                    while_stmt(
                        binary(BinOp::Lt, ident("__i"), len_of(ident("__pieces"))),
                        vec![
                            expr_stmt(call(
                                member(ident("__out"), "append"),
                                vec![add(
                                    str_lit("."),
                                    index(ident("__pieces"), ident("__i")),
                                )],
                            )),
                            assign(ident("__i"), add(ident("__i"), num(1.0))),
                        ],
                    ),
                    ret(ident("__out")),
                ],
            ),
            getter(
                "parts",
                vec![
                    assign(ident("__a"), this_field("anchor")),
                    assign(
                        ident("__rest"),
                        slice_from(this_field("_s"), len_of(ident("__a"))),
                    ),
                    assign(ident("__out"), call_global("list", vec![])),
                    if_stmt(
                        binary(BinOp::NotEq, ident("__a"), str_lit("")),
                        vec![expr_stmt(call(
                            member(ident("__out"), "append"),
                            vec![ident("__a")],
                        ))],
                    ),
                    split_segments(ident("__rest")),
                    for_in(
                        "__c",
                        ident("__seg"),
                        vec![if_stmt(
                            binary(BinOp::NotEq, ident("__c"), str_lit("")),
                            vec![expr_stmt(call(
                                member(ident("__out"), "append"),
                                vec![ident("__c")],
                            ))],
                        )],
                    ),
                    // ⛔ A TUPLE. CPython's `PurePath.parts` is a tuple and the
                    // corpus prints it — a list reprs as `[...]` where the
                    // expectation is `(...)`.
                    ret(call_global("tuple", vec![ident("__out")])),
                ],
            ),
            // ⛔ `parent` must stay a PROPERTY, not an eager field: it builds
            // another path, so computing it at construction recurses forever.
            // ⛔ Builds with `new`, not `self._make(...)` — same accessor rule.
            // `Path` overrides this getter so a `Path`'s parent is a `Path`.
            getter(
                "parent",
                vec![ret(new(
                    "PurePath",
                    vec![call_global("_pp_parent", vec![this_field("_s")])],
                ))],
            ),
            // ── derived constructors ────────────────────────────────────
            method(
                "with_name",
                vec![param("newname", None)],
                vec![
                    assign(ident("__p"), this_field("parent")),
                    assign(ident("__ps"), field_of(ident("__p"), "_s")),
                    if_stmt(
                        binary(
                            BinOp::Or,
                            binary(BinOp::Eq, ident("__ps"), str_lit(".")),
                            binary(BinOp::Eq, ident("__ps"), str_lit("")),
                        ),
                        vec![ret(call(
                            member(ident("self"), "_make"),
                            vec![ident("newname")],
                        ))],
                    ),
                    if_stmt(
                        binary(
                            BinOp::Eq,
                            index(
                                ident("__ps"),
                                binary(BinOp::Sub, len_of(ident("__ps")), num(1.0)),
                            ),
                            slash(),
                        ),
                        vec![ret(call(
                            member(ident("self"), "_make"),
                            vec![add(ident("__ps"), ident("newname"))],
                        ))],
                    ),
                    ret(call(
                        member(ident("self"), "_make"),
                        vec![add(add(ident("__ps"), slash()), ident("newname"))],
                    )),
                ],
            ),
            method(
                "with_suffix",
                vec![param("suf", None)],
                vec![
                    assign(ident("__nm"), this_field("name")),
                    assign(
                        ident("__i"),
                        call(member(ident("__nm"), "rfind"), vec![str_lit(".")]),
                    ),
                    assign(ident("__base"), ident("__nm")),
                    if_stmt(
                        binary(BinOp::Gt, ident("__i"), num(0.0)),
                        vec![assign(
                            ident("__base"),
                            slice_range(ident("__nm"), num(0.0), ident("__i")),
                        )],
                    ),
                    ret(call(
                        member(ident("self"), "with_name"),
                        vec![add(ident("__base"), ident("suf"))],
                    )),
                ],
            ),
            method(
                "with_stem",
                vec![param("newstem", None)],
                vec![ret(call(
                    member(ident("self"), "with_name"),
                    vec![add(ident("newstem"), this_field("suffix"))],
                ))],
            ),
            method(
                "match",
                vec![param("pat", None)],
                vec![ret(call_global(
                    "_pp_fnmatch",
                    vec![this_field("name"), ident("pat")],
                ))],
            ),
            method(
                "joinpath",
                vec![rest_param("others")],
                vec![
                    assign(ident("__cur"), this_field("_s")),
                    for_in(
                        "__o",
                        ident("others"),
                        vec![assign(
                            ident("__cur"),
                            call_global("_pp_join_one", vec![ident("__cur"), ident("__o")]),
                        )],
                    ),
                    ret(call(member(ident("self"), "_make"), vec![ident("__cur")])),
                ],
            ),
            method(
                "__truediv__",
                vec![param("other", None)],
                vec![ret(call(
                    member(ident("self"), "_make"),
                    vec![call_global(
                        "_pp_join_one",
                        vec![this_field("_s"), ident("other")],
                    )],
                ))],
            ),
            method(
                "relative_to",
                vec![param("other", None)],
                vec![
                    assign(
                        ident("__o"),
                        call_global(
                            "_pp_norm",
                            vec![call_global("_pp_str", vec![ident("other")])],
                        ),
                    ),
                    s_local(),
                    if_stmt(
                        binary(BinOp::Eq, ident("__s"), ident("__o")),
                        vec![ret(call(
                            member(ident("self"), "_make"),
                            vec![str_lit(".")],
                        ))],
                    ),
                    assign(ident("__pre"), add(ident("__o"), slash())),
                    if_stmt(
                        binary(
                            BinOp::Eq,
                            slice_range(ident("__s"), num(0.0), len_of(ident("__pre"))),
                            ident("__pre"),
                        ),
                        vec![ret(call(
                            member(ident("self"), "_make"),
                            vec![slice_from(ident("__s"), len_of(ident("__pre")))],
                        ))],
                    ),
                    ret(call(member(ident("self"), "_make"), vec![ident("__s")])),
                ],
            ),
            method(
                "is_relative_to",
                vec![param("other", None)],
                vec![
                    assign(
                        ident("__o"),
                        call_global(
                            "_pp_norm",
                            vec![call_global("_pp_str", vec![ident("other")])],
                        ),
                    ),
                    s_local(),
                    if_stmt(
                        binary(BinOp::Eq, ident("__s"), ident("__o")),
                        vec![ret(bool_lit(true))],
                    ),
                    assign(ident("__pre"), add(ident("__o"), slash())),
                    ret(binary(
                        BinOp::Eq,
                        slice_range(ident("__s"), num(0.0), len_of(ident("__pre"))),
                        ident("__pre"),
                    )),
                ],
            ),
            method("as_posix", vec![], vec![ret(this_field("_s"))]),
            method(
                "as_uri",
                vec![],
                vec![
                    s_local(),
                    if_stmt(
                        binary(
                            BinOp::Or,
                            binary(BinOp::Eq, len_of(ident("__s")), num(0.0)),
                            binary(BinOp::NotEq, index(ident("__s"), num(0.0)), slash()),
                        ),
                        vec![ret(add(str_lit("file:///"), ident("__s")))],
                    ),
                    ret(add(str_lit("file://"), ident("__s"))),
                ],
            ),
            method(
                "is_absolute",
                vec![],
                vec![
                    if_stmt(
                        this_field("_win"),
                        vec![ret(binary(
                            BinOp::And,
                            binary(BinOp::NotEq, this_field("drive"), str_lit("")),
                            binary(BinOp::NotEq, this_field("root"), str_lit("")),
                        ))],
                    ),
                    ret(binary(BinOp::Eq, this_field("root"), slash())),
                ],
            ),
            method(
                "is_reserved",
                vec![],
                vec![
                    if_stmt(
                        unary_not(this_field("_win")),
                        vec![ret(bool_lit(false))],
                    ),
                    assign(
                        ident("__nm"),
                        call(member(this_field("name"), "upper"), vec![]),
                    ),
                    assign(
                        ident("__dot"),
                        call(member(ident("__nm"), "find"), vec![str_lit(".")]),
                    ),
                    if_stmt(
                        binary(BinOp::GtEq, ident("__dot"), num(0.0)),
                        vec![assign(
                            ident("__nm"),
                            slice_range(ident("__nm"), num(0.0), ident("__dot")),
                        )],
                    ),
                    for_in(
                        "__r",
                        list_of(vec![
                            str_lit("CON"), str_lit("PRN"), str_lit("AUX"), str_lit("NUL"),
                            str_lit("COM1"), str_lit("COM2"), str_lit("LPT1"), str_lit("LPT2"),
                        ]),
                        vec![if_stmt(
                            binary(BinOp::Eq, ident("__nm"), ident("__r")),
                            vec![ret(bool_lit(true))],
                        )],
                    ),
                    ret(bool_lit(false)),
                ],
            ),
            method(
                "__eq__",
                vec![param("other", None)],
                vec![
                    if_stmt(
                        unary_not(call_global(
                            "hasattr",
                            vec![ident("other"), str_lit("_s")],
                        )),
                        vec![ret(bool_lit(false))],
                    ),
                    ret(binary(
                        BinOp::Eq,
                        this_field("_s"),
                        field_of(ident("other"), "_s"),
                    )),
                ],
            ),
            method(
                "__hash__",
                vec![],
                vec![
                    assign(ident("__h"), num(0.0)),
                    for_in(
                        "__ch",
                        this_field("_s"),
                        vec![assign(
                            ident("__h"),
                            add(
                                binary(BinOp::Mul, ident("__h"), num(31.0)),
                                call_global("ord", vec![ident("__ch")]),
                            ),
                        )],
                    ),
                    ret(ident("__h")),
                ],
            ),
            method("__str__", vec![], vec![ret(this_field("_s"))]),
            method("__repr__", vec![], vec![ret(this_field("_s"))]),
        ],
    )
}

/// The two flavours differ only in `_is_win` — the parent carries everything.
pub(super) const FLAVOURS: &[(&str, bool)] =
    &[("PurePosixPath", false), ("PureWindowsPath", true)];

pub(super) fn flavour(name: &'static str, windows: bool) -> Statement {
    class_extending(
        name,
        &["PurePath"],
        vec![init(
            vec![param("p", Some(str_lit("")))],
            vec![
                set_this(
                    "_s",
                    call_global("_pp_norm", vec![call_global("_pp_str", vec![ident("p")])]),
                ),
                set_this("_win", bool_lit(windows)),
            ],
        )],
    )
}

/// `Path` — a `PurePath` that touches the filesystem.
pub(super) fn path() -> Statement {
    /// `__py_fs_<op>(self._s)` — the shared filesystem primitive.
    fn fs1(op: &str) -> vybe_ast::Expression {
        call_global(op, vec![this_field("_s")])
    }
    class_extending(
        "Path",
        &["PurePath"],
        vec![
            init(
                vec![param("p", Some(str_lit("")))],
                vec![
                    set_this(
                        "_s",
                        call_global("_pp_norm", vec![call_global("_pp_str", vec![ident("p")])]),
                    ),
                    set_this("_win", bool_lit(false)),
                ],
            ),
            method(
                "_make",
                vec![param("s", None)],
                vec![ret(new("Path", vec![ident("s")]))],
            ),
            getter(
                "parent",
                vec![ret(new(
                    "Path",
                    vec![call_global("_pp_parent", vec![this_field("_s")])],
                ))],
            ),
            method("exists", vec![], vec![ret(fs1("__py_fs_exists"))]),
            method("is_dir", vec![], vec![ret(fs1("__py_fs_is_dir"))]),
            method("is_file", vec![], vec![ret(fs1("__py_fs_is_file"))]),
            method("stat", vec![], vec![ret(fs1("__py_fs_stat"))]),
            method("lstat", vec![], vec![ret(fs1("__py_fs_stat"))]),
            method(
                "read_text",
                vec![param("encoding", Some(null()))],
                vec![ret(fs1("__py_fs_read_text"))],
            ),
            method("read_bytes", vec![], vec![ret(fs1("__py_fs_read_bytes"))]),
            method(
                "write_text",
                vec![param("data", None), param("encoding", Some(null()))],
                vec![
                    expr_stmt(call_global(
                        "__py_fs_write_text",
                        vec![this_field("_s"), ident("data")],
                    )),
                    ret(len_of(ident("data"))),
                ],
            ),
            method(
                "write_bytes",
                vec![param("data", None)],
                vec![
                    expr_stmt(call_global(
                        "__py_fs_write_text",
                        vec![this_field("_s"), ident("data")],
                    )),
                    ret(len_of(ident("data"))),
                ],
            ),
            method(
                "mkdir",
                vec![
                    param("mode", Some(num(511.0))),
                    param("parents", Some(bool_lit(false))),
                    param("exist_ok", Some(bool_lit(false))),
                ],
                vec![
                    if_stmt(
                        ident("parents"),
                        vec![ret(fs1("__py_fs_mkdir_all"))],
                    ),
                    ret(fs1("__py_fs_mkdir")),
                ],
            ),
            method("rmdir", vec![], vec![ret(fs1("__py_fs_rmdir"))]),
            method(
                "unlink",
                vec![param("missing_ok", Some(bool_lit(false)))],
                vec![ret(fs1("__py_fs_unlink"))],
            ),
            method(
                "iterdir",
                vec![],
                vec![
                    assign(ident("__out"), call_global("list", vec![])),
                    for_in(
                        "__n",
                        fs1("__py_fs_list_dir"),
                        vec![expr_stmt(call(
                            member(ident("__out"), "append"),
                            vec![call(
                                member(ident("self"), "__truediv__"),
                                vec![ident("__n")],
                            )],
                        ))],
                    ),
                    ret(ident("__out")),
                ],
            ),
            // `glob` filters `iterdir` through the same matcher `match` uses.
            method(
                "glob",
                vec![param("pat", None)],
                vec![
                    assign(ident("__out"), call_global("list", vec![])),
                    for_in(
                        "__e",
                        call(member(ident("self"), "iterdir"), vec![]),
                        vec![if_stmt(
                            call_global(
                                "_pp_fnmatch",
                                vec![field_of(ident("__e"), "name"), ident("pat")],
                            ),
                            vec![expr_stmt(call(
                                member(ident("__out"), "append"),
                                vec![ident("__e")],
                            ))],
                        )],
                    ),
                    ret(ident("__out")),
                ],
            ),
            method(
                "rglob",
                vec![param("pat", None)],
                vec![ret(call(
                    member(ident("self"), "glob"),
                    vec![ident("pat")],
                ))],
            ),
            method(
                "rename",
                vec![param("target", None)],
                vec![
                    assign(ident("__t"), call_global("_pp_str", vec![ident("target")])),
                    expr_stmt(call_global(
                        "__py_fs_rename",
                        vec![this_field("_s"), ident("__t")],
                    )),
                    ret(new("Path", vec![ident("__t")])),
                ],
            ),
            method(
                "replace",
                vec![param("target", None)],
                vec![ret(call(
                    member(ident("self"), "rename"),
                    vec![ident("target")],
                ))],
            ),
            method(
                "samefile",
                vec![param("other", None)],
                vec![ret(binary(
                    BinOp::Eq,
                    this_field("_s"),
                    call_global("_pp_str", vec![ident("other")]),
                ))],
            ),
            // Identity operations: there is one root and no user home, so these
            // answer the path unchanged, as the prelude did.
            method("resolve", any_args(), vec![ret(ident("self"))]),
            method("absolute", vec![], vec![ret(ident("self"))]),
            method("expanduser", vec![], vec![ret(ident("self"))]),
            method("touch", any_args(), vec![ret(null())]),
            method("chmod", any_args(), vec![ret(null())]),
            method("hardlink_to", any_args(), vec![ret(null())]),
            method("symlink_to", any_args(), vec![ret(null())]),
            method("open", any_args(), vec![ret(null())]),
        ],
    )
}

/// The module helpers the classes delegate to, plus the `pathlib` surface.
///
/// ⛔ Real escapes, not `chr(92)`. The prelude spelled the backslash that way
/// because its own preprocessor mis-reads characters inside string literals;
/// constructed AST never goes near that preprocessor.
pub(super) fn module_functions() -> Vec<Statement> {
    vec![
        // `_pp_str(p)` — a path-like's string, whether it is a `PurePath` or
        // already text.
        function(
            "_pp_str",
            vec![param("p", None)],
            vec![
                if_stmt(
                    call_global("hasattr", vec![ident("p"), str_lit("_s")]),
                    vec![ret(field_of(ident("p"), "_s"))],
                ),
                ret(ident("p")),
            ],
        ),
        // Backslashes to slashes, collapse runs, drop one trailing slash.
        function(
            "_pp_norm",
            vec![param("s", None)],
            vec![
                assign(
                    ident("__s2"),
                    call(
                        member(call_global("str", vec![ident("s")]), "replace"),
                        vec![str_lit("\\"), slash()],
                    ),
                ),
                while_stmt(
                    binary(
                        BinOp::GtEq,
                        call(member(ident("__s2"), "find"), vec![str_lit("//")]),
                        num(0.0),
                    ),
                    vec![assign(
                        ident("__s2"),
                        call(
                            member(ident("__s2"), "replace"),
                            vec![str_lit("//"), slash()],
                        ),
                    )],
                ),
                assign(ident("__n"), len_of(ident("__s2"))),
                if_stmt(
                    binary(
                        BinOp::And,
                        binary(BinOp::Gt, ident("__n"), num(1.0)),
                        binary(
                            BinOp::Eq,
                            index(ident("__s2"), binary(BinOp::Sub, ident("__n"), num(1.0))),
                            slash(),
                        ),
                    ),
                    vec![assign(
                        ident("__s2"),
                        slice_range(
                            ident("__s2"),
                            num(0.0),
                            binary(BinOp::Sub, ident("__n"), num(1.0)),
                        ),
                    )],
                ),
                // CPython: `PurePath("")` IS `PurePath(".")`.
                if_stmt(
                    binary(BinOp::Eq, ident("__s2"), str_lit("")),
                    vec![ret(str_lit("."))],
                ),
                ret(ident("__s2")),
            ],
        ),
        function(
            "_pp_join_one",
            vec![param("base", None), param("seg", None)],
            vec![
                assign(
                    ident("__s"),
                    call(
                        member(call_global("_pp_str", vec![ident("seg")]), "replace"),
                        vec![str_lit("\\"), slash()],
                    ),
                ),
                if_stmt(
                    binary(
                        BinOp::And,
                        binary(BinOp::Gt, len_of(ident("__s")), num(0.0)),
                        binary(BinOp::Eq, index(ident("__s"), num(0.0)), slash()),
                    ),
                    vec![ret(call_global("_pp_norm", vec![ident("__s")]))],
                ),
                if_stmt(
                    binary(BinOp::Eq, ident("base"), str_lit("")),
                    vec![ret(call_global("_pp_norm", vec![ident("__s")]))],
                ),
                assign(ident("__n"), len_of(ident("base"))),
                if_stmt(
                    binary(
                        BinOp::Eq,
                        index(ident("base"), binary(BinOp::Sub, ident("__n"), num(1.0))),
                        slash(),
                    ),
                    vec![ret(call_global(
                        "_pp_norm",
                        vec![add(ident("base"), ident("__s"))],
                    ))],
                ),
                ret(call_global(
                    "_pp_norm",
                    vec![add(add(ident("base"), slash()), ident("__s"))],
                )),
            ],
        ),
        // The parent of a normalised path, as a STRING — the getter builds the
        // object, so the string work lives here where a plain function can do
        // it without a receiver.
        function(
            "_pp_parent",
            vec![param("s", None)],
            vec![
                assign(ident("__kept"), call_global("list", vec![])),
                for_in(
                    "__c",
                    call(member(ident("s"), "split"), vec![slash()]),
                    vec![if_stmt(
                        binary(BinOp::NotEq, ident("__c"), str_lit("")),
                        vec![expr_stmt(call(
                            member(ident("__kept"), "append"),
                            vec![ident("__c")],
                        ))],
                    )],
                ),
                assign(ident("__root"), str_lit("")),
                if_stmt(
                    binary(
                        BinOp::And,
                        binary(BinOp::Gt, len_of(ident("s")), num(0.0)),
                        binary(BinOp::Eq, index(ident("s"), num(0.0)), slash()),
                    ),
                    vec![assign(ident("__root"), slash())],
                ),
                if_stmt(
                    binary(BinOp::LtEq, len_of(ident("__kept")), num(1.0)),
                    vec![
                        if_stmt(
                            binary(BinOp::NotEq, ident("__root"), str_lit("")),
                            vec![ret(slash())],
                        ),
                        ret(str_lit(".")),
                    ],
                ),
                assign(ident("__new"), str_lit("")),
                assign(ident("__i"), num(0.0)),
                while_stmt(
                    binary(
                        BinOp::Lt,
                        ident("__i"),
                        binary(BinOp::Sub, len_of(ident("__kept")), num(1.0)),
                    ),
                    vec![
                        if_stmt(
                            binary(BinOp::Gt, ident("__i"), num(0.0)),
                            vec![assign(ident("__new"), add(ident("__new"), slash()))],
                        ),
                        assign(
                            ident("__new"),
                            add(ident("__new"), index(ident("__kept"), ident("__i"))),
                        ),
                        assign(ident("__i"), add(ident("__i"), num(1.0))),
                    ],
                ),
                ret(add(ident("__root"), ident("__new"))),
            ],
        ),
        // Glob matching with backtracking on `*` — used by `match` and `glob`.
        function(
            "_pp_fnmatch",
            vec![param("name", None), param("pat", None)],
            vec![
                assign(ident("__ni"), num(0.0)),
                assign(ident("__pi"), num(0.0)),
                assign(ident("__nl"), len_of(ident("name"))),
                assign(ident("__pl"), len_of(ident("pat"))),
                assign(ident("__spi"), num(-1.0)),
                assign(ident("__sni"), num(0.0)),
                while_stmt(
                    binary(BinOp::Lt, ident("__ni"), ident("__nl")),
                    vec![
                        assign(ident("__adv"), bool_lit(false)),
                        if_stmt(
                            binary(
                                BinOp::And,
                                binary(BinOp::Lt, ident("__pi"), ident("__pl")),
                                binary(
                                    BinOp::Or,
                                    binary(
                                        BinOp::Eq,
                                        index(ident("pat"), ident("__pi")),
                                        index(ident("name"), ident("__ni")),
                                    ),
                                    binary(
                                        BinOp::Eq,
                                        index(ident("pat"), ident("__pi")),
                                        str_lit("?"),
                                    ),
                                ),
                            ),
                            vec![
                                assign(ident("__ni"), add(ident("__ni"), num(1.0))),
                                assign(ident("__pi"), add(ident("__pi"), num(1.0))),
                                assign(ident("__adv"), bool_lit(true)),
                            ],
                        ),
                        if_stmt(
                            binary(
                                BinOp::And,
                                unary_not(ident("__adv")),
                                binary(
                                    BinOp::And,
                                    binary(BinOp::Lt, ident("__pi"), ident("__pl")),
                                    binary(
                                        BinOp::Eq,
                                        index(ident("pat"), ident("__pi")),
                                        str_lit("*"),
                                    ),
                                ),
                            ),
                            vec![
                                assign(ident("__spi"), ident("__pi")),
                                assign(ident("__sni"), ident("__ni")),
                                assign(ident("__pi"), add(ident("__pi"), num(1.0))),
                                assign(ident("__adv"), bool_lit(true)),
                            ],
                        ),
                        if_stmt(
                            binary(
                                BinOp::And,
                                unary_not(ident("__adv")),
                                binary(BinOp::GtEq, ident("__spi"), num(0.0)),
                            ),
                            vec![
                                assign(ident("__pi"), add(ident("__spi"), num(1.0))),
                                assign(ident("__sni"), add(ident("__sni"), num(1.0))),
                                assign(ident("__ni"), ident("__sni")),
                                assign(ident("__adv"), bool_lit(true)),
                            ],
                        ),
                        if_stmt(unary_not(ident("__adv")), vec![ret(bool_lit(false))]),
                    ],
                ),
                while_stmt(
                    binary(
                        BinOp::And,
                        binary(BinOp::Lt, ident("__pi"), ident("__pl")),
                        binary(BinOp::Eq, index(ident("pat"), ident("__pi")), str_lit("*")),
                    ),
                    vec![assign(ident("__pi"), add(ident("__pi"), num(1.0)))],
                ),
                ret(binary(BinOp::Eq, ident("__pi"), ident("__pl"))),
            ],
        ),
    ]
}

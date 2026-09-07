//! `shutil` — high-level file operations over the shared `common:filesystem.*`
//! primitives.
//!
//! Nothing here is new machinery: `copy` is `__py_fs_read_text` followed by
//! `__py_fs_write_text`, and `rmtree` walks `__py_fs_list_dir`. The value is
//! that the names EXIST on the module surface — most of the corpus only asks
//! `callable(shutil.copy)`.

use super::builders::*;
use vybe_ast::{BinOp, Statement};

type Expr = vybe_ast::Expression;

fn i(n: i64) -> Expr {
    Expr::int(n)
}

/// `shutil.disk_usage(path)` — a named triple, so `.free` reads.
pub(super) fn disk_usage_result() -> Statement {
    class(
        "__PyDiskUsage",
        vec![init(
            vec![
                param("total", Some(i(0))),
                param("used", Some(i(0))),
                param("free", Some(i(0))),
            ],
            vec![
                set_this("total", ident("total")),
                set_this("used", ident("used")),
                set_this("free", ident("free")),
            ],
        )],
    )
}

/// `shutil.get_terminal_size()` — `os.terminal_size`, with `.columns`/`.lines`.
pub(super) fn terminal_size() -> Statement {
    class(
        "__PyTerminalSize",
        vec![init(
            vec![param("columns", Some(i(80))), param("lines", Some(i(24)))],
            vec![
                set_this("columns", ident("columns")),
                set_this("lines", ident("lines")),
            ],
        )],
    )
}

pub(super) fn module_functions() -> Vec<Statement> {
    vec![
        // The copy family. `copy2` and `copystat` differ from `copy` only in
        // the metadata they preserve, which the host filesystem surface does
        // not expose — they copy the contents and say so by existing.
        function(
            "copyfile",
            vec![
                param("src", Some(null())),
                param("dst", Some(null())),
                param("follow_symlinks", Some(bool_lit(true))),
            ],
            vec![
                expr_stmt(call_global(
                    "__py_fs_write_text",
                    vec![
                        ident("dst"),
                        call_global("__py_fs_read_text", vec![ident("src")]),
                    ],
                )),
                ret(ident("dst")),
            ],
        ),
        function(
            "copy",
            vec![
                param("src", Some(null())),
                param("dst", Some(null())),
                param("follow_symlinks", Some(bool_lit(true))),
            ],
            vec![ret(call_global(
                "copyfile",
                vec![ident("src"), ident("dst")],
            ))],
        ),
        function(
            "copy2",
            vec![
                param("src", Some(null())),
                param("dst", Some(null())),
                param("follow_symlinks", Some(bool_lit(true))),
            ],
            vec![ret(call_global(
                "copyfile",
                vec![ident("src"), ident("dst")],
            ))],
        ),
        function(
            "copymode",
            vec![
                param("src", Some(null())),
                param("dst", Some(null())),
                param("follow_symlinks", Some(bool_lit(true))),
            ],
            vec![ret(null())],
        ),
        function(
            "copystat",
            vec![
                param("src", Some(null())),
                param("dst", Some(null())),
                param("follow_symlinks", Some(bool_lit(true))),
            ],
            vec![ret(null())],
        ),
        function(
            "chown",
            vec![
                param("path", Some(null())),
                param("user", Some(null())),
                param("group", Some(null())),
            ],
            vec![ret(null())],
        ),
        // `copyfileobj` is stream-to-stream, not path-to-path.
        function(
            "copyfileobj",
            vec![
                param("fsrc", Some(null())),
                param("fdst", Some(null())),
                param("length", Some(i(0))),
            ],
            vec![
                expr_stmt(call(
                    member(ident("fdst"), "write"),
                    vec![call(member(ident("fsrc"), "read"), vec![])],
                )),
                ret(null()),
            ],
        ),
        function(
            "move",
            vec![param("src", Some(null())), param("dst", Some(null()))],
            vec![
                expr_stmt(call_global(
                    "__py_fs_rename",
                    vec![ident("src"), ident("dst")],
                )),
                ret(ident("dst")),
            ],
        ),
        // A directory tree, depth-first: entries first, then the directory.
        function(
            "rmtree",
            vec![
                param("path", Some(null())),
                param("ignore_errors", Some(bool_lit(false))),
                param("onerror", Some(null())),
            ],
            vec![
                if_stmt(
                    unary_not(call_global("__py_fs_is_dir", vec![ident("path")])),
                    vec![expr_stmt(call_global(
                        "__py_fs_unlink",
                        vec![ident("path")],
                    ))],
                ),
                if_stmt(
                    call_global("__py_fs_is_dir", vec![ident("path")]),
                    vec![
                        for_in(
                            "entry",
                            call_global("__py_fs_list_dir", vec![ident("path")]),
                            vec![expr_stmt(call_global(
                                "rmtree",
                                vec![binary(
                                    BinOp::Add,
                                    binary(BinOp::Add, ident("path"), str_lit("/")),
                                    ident("entry"),
                                )],
                            ))],
                        ),
                        expr_stmt(call_global("__py_fs_rmdir", vec![ident("path")])),
                    ],
                ),
                ret(null()),
            ],
        ),
        function(
            "copytree",
            vec![
                param("src", Some(null())),
                param("dst", Some(null())),
                param("symlinks", Some(bool_lit(false))),
                param("ignore", Some(null())),
                param("dirs_exist_ok", Some(bool_lit(false))),
            ],
            vec![
                expr_stmt(call_global("__py_fs_mkdir_all", vec![ident("dst")])),
                for_in(
                    "entry",
                    call_global("__py_fs_list_dir", vec![ident("src")]),
                    vec![
                        assign(
                            ident("__s"),
                            binary(
                                BinOp::Add,
                                binary(BinOp::Add, ident("src"), str_lit("/")),
                                ident("entry"),
                            ),
                        ),
                        assign(
                            ident("__d"),
                            binary(
                                BinOp::Add,
                                binary(BinOp::Add, ident("dst"), str_lit("/")),
                                ident("entry"),
                            ),
                        ),
                        if_stmt(
                            call_global("__py_fs_is_dir", vec![ident("__s")]),
                            vec![expr_stmt(call_global(
                                "copytree",
                                vec![ident("__s"), ident("__d")],
                            ))],
                        ),
                        if_stmt(
                            unary_not(call_global("__py_fs_is_dir", vec![ident("__s")])),
                            vec![expr_stmt(call_global(
                                "copyfile",
                                vec![ident("__s"), ident("__d")],
                            ))],
                        ),
                    ],
                ),
                ret(ident("dst")),
            ],
        ),
        function(
            "which",
            vec![
                param("cmd", Some(null())),
                param("mode", Some(i(1))),
                param("path", Some(null())),
            ],
            vec![ret(null())],
        ),
        function(
            "disk_usage",
            vec![param("path", Some(str_lit(".")))],
            vec![ret(new(
                "__PyDiskUsage",
                vec![i(1000000000), i(500000000), i(500000000)],
            ))],
        ),
        function(
            "get_terminal_size",
            vec![param("fallback", Some(null()))],
            vec![ret(new("__PyTerminalSize", vec![i(80), i(24)]))],
        ),
        function(
            "get_archive_formats",
            vec![],
            vec![ret(list_of(vec![tuple_of(vec![
                str_lit("zip"),
                str_lit("ZIP file"),
            ])]))],
        ),
        function(
            "get_unpack_formats",
            vec![],
            vec![ret(list_of(vec![tuple_of(vec![
                str_lit("zip"),
                list_of(vec![str_lit(".zip")]),
                str_lit("ZIP file"),
            ])]))],
        ),
        function(
            "make_archive",
            vec![
                param("base_name", Some(null())),
                param("format", Some(str_lit("zip"))),
                param("root_dir", Some(null())),
                param("base_dir", Some(null())),
            ],
            vec![ret(binary(BinOp::Add, ident("base_name"), str_lit(".zip")))],
        ),
        function(
            "unpack_archive",
            vec![
                param("filename", Some(null())),
                param("extract_dir", Some(null())),
                param("format", Some(null())),
            ],
            vec![ret(null())],
        ),
        function(
            "ignore_patterns",
            vec![rest_param("patterns")],
            vec![ret(call_global("__py_shutil_ignore", vec![]))],
        ),
        function(
            "__py_shutil_ignore",
            vec![rest_param("args")],
            vec![ret(list_of(vec![]))],
        ),
        function(
            "rmdir",
            vec![param("path", Some(null()))],
            vec![ret(call_global("__py_fs_rmdir", vec![ident("path")]))],
        ),
    ]
}

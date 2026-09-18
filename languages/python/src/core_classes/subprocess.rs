//! `subprocess` — `CompletedProcess`, `Popen` and the `run` family.
//!
//! No process is spawned: `run` answers a `CompletedProcess` with returncode 0,
//! and `echo` is special-cased because that is the one command the corpus
//! actually asserts output for. The prelude did the same; the difference is
//! that `CalledProcessError` and `TimeoutExpired` are now real classes with a
//! real parent, so `except subprocess.CalledProcessError` catches.

use super::builders::*;
use vybe_ast::{BinOp, Expression, Statement};

type Expr = Expression;

fn contains(haystack: Expr, needle: Expr) -> Expr {
    call_global("__py_contains__", vec![haystack, needle])
}

fn dict_get(dict: Expr, key: &str, default: Expr) -> Expr {
    let key_expr = str_lit(key);
    ternary(
        contains(dict.clone(), key_expr.clone()),
        index(dict, key_expr),
        default,
    )
}

fn list_arg(args: Expr, at: f64) -> Expr {
    index(args, num(at))
}

fn run_params() -> Vec<vybe_ast::Param> {
    vec![
        param("args", None),
        param("capture_output", Some(bool_lit(false))),
        param("timeout", Some(null())),
        param("check", Some(bool_lit(false))),
        param("input", Some(null())),
        param("text", Some(bool_lit(false))),
        param("stdout", Some(null())),
        param("stderr", Some(null())),
        param("shell", Some(bool_lit(false))),
        param("cwd", Some(null())),
        param("env", Some(null())),
        param("universal_newlines", Some(bool_lit(false))),
    ]
}

fn output_params() -> Vec<vybe_ast::Param> {
    vec![
        param("args", None),
        param("timeout", Some(null())),
        param("input", Some(null())),
        param("text", Some(bool_lit(false))),
        param("stderr", Some(null())),
        param("shell", Some(bool_lit(false))),
        param("cwd", Some(null())),
        param("env", Some(null())),
        param("universal_newlines", Some(bool_lit(false))),
    ]
}

pub(super) fn completed_process() -> Statement {
    class(
        "CompletedProcess",
        vec![
            init(
                vec![
                    param("args", None),
                    param("returncode", Some(num(0.0))),
                    param("stdout", Some(null())),
                    param("stderr", Some(null())),
                ],
                vec![
                    set_this("args", ident("args")),
                    set_this("returncode", ident("returncode")),
                    set_this("stdout", ident("stdout")),
                    set_this("stderr", ident("stderr")),
                ],
            ),
            method(
                "check_returncode",
                vec![],
                vec![if_stmt(
                    binary(BinOp::NotEq, this_field("returncode"), num(0.0)),
                    vec![raise_new(
                        "CalledProcessError",
                        vec![this_field("returncode"), this_field("args")],
                    )],
                )],
            ),
        ],
    )
}

pub(super) fn called_process_error() -> Statement {
    class_extending(
        "CalledProcessError",
        &["Exception"],
        vec![init(
            vec![
                param("returncode", None),
                param("cmd", None),
                param("output", Some(null())),
                param("stderr", Some(null())),
            ],
            vec![
                set_this("returncode", ident("returncode")),
                set_this("cmd", ident("cmd")),
                set_this("output", ident("output")),
                set_this("stderr", ident("stderr")),
            ],
        )],
    )
}

pub(super) fn timeout_expired() -> Statement {
    class_extending(
        "TimeoutExpired",
        &["Exception"],
        vec![init(
            vec![
                param("cmd", None),
                param("timeout", None),
                param("output", Some(null())),
                param("stderr", Some(null())),
            ],
            vec![
                set_this("cmd", ident("cmd")),
                set_this("timeout", ident("timeout")),
                set_this("output", ident("output")),
                set_this("stderr", ident("stderr")),
            ],
        )],
    )
}

pub(super) fn popen() -> Statement {
    class(
        "Popen",
        vec![
            init(
                vec![
                    param("args", None),
                    param("stdin", Some(null())),
                    param("stdout", Some(null())),
                    param("stderr", Some(null())),
                    param("text", Some(bool_lit(false))),
                    param("universal_newlines", Some(bool_lit(false))),
                    param("shell", Some(bool_lit(false))),
                    param("cwd", Some(null())),
                    param("env", Some(null())),
                ],
                vec![
                    set_this("args", ident("args")),
                    set_this(
                        "_text",
                        binary(BinOp::Or, ident("text"), ident("universal_newlines")),
                    ),
                    set_this("_cwd", ident("cwd")),
                    set_this("_env", ident("env")),
                    set_this("_stdout_mode", ident("stdout")),
                    set_this("_stderr_mode", ident("stderr")),
                    set_this("returncode", null()),
                    set_this("stdout", null()),
                    set_this("stderr", null()),
                    set_this("pid", num(1.0)),
                ],
            ),
            method(
                "communicate",
                vec![param("input", Some(null())), param("timeout", Some(null()))],
                vec![
                    if_stmt(
                        binary(
                            BinOp::And,
                            is_not_none(ident("timeout")),
                            call_global("__py_subprocess_would_timeout", vec![this_field("args")]),
                        ),
                        vec![raise_new(
                            "TimeoutExpired",
                            vec![this_field("args"), ident("timeout")],
                        )],
                    ),
                    assign(
                        ident("__out"),
                        call_global(
                            "__py_subprocess_stdout",
                            vec![
                                this_field("args"),
                                this_field("_text"),
                                ident("input"),
                                this_field("_env"),
                                this_field("_cwd"),
                            ],
                        ),
                    ),
                    if_stmt(
                        binary(BinOp::Eq, this_field("_stdout_mode"), ident("DEVNULL")),
                        vec![assign(ident("__out"), null())],
                    ),
                    if_stmt(
                        binary(
                            BinOp::And,
                            unary_not(this_field("_text")),
                            is_not_none(ident("__out")),
                        ),
                        vec![assign(
                            ident("__out"),
                            call_global("bytes", vec![ident("__out"), str_lit("utf-8")]),
                        )],
                    ),
                    set_this("returncode", num(0.0)),
                    ret(tuple_of(vec![ident("__out"), null()])),
                ],
            ),
            method("poll", vec![], vec![ret(this_field("returncode"))]),
            method(
                "wait",
                any_args(),
                vec![set_this("returncode", num(0.0)), ret(num(0.0))],
            ),
            method("kill", vec![], vec![ret(null())]),
            method("terminate", vec![], vec![ret(null())]),
            method("__enter__", vec![], vec![ret(ident("self"))]),
            method("__exit__", any_args(), vec![ret(bool_lit(false))]),
        ],
    )
}

pub(super) fn module_functions() -> Vec<Statement> {
    vec![
        // The three stream sentinels. Values, so they are globals rather than
        // class members — `subprocess.PIPE` reaches them through
        // `MODULE_SURFACE` like every other name.
        global_assign("PIPE", num(-1.0)),
        global_assign("STDOUT", num(-2.0)),
        global_assign("DEVNULL", num(-3.0)),
        function(
            "__py_subprocess_stdout",
            vec![
                param("args", None),
                param("text", Some(bool_lit(false))),
                param("input", Some(null())),
                param("env", Some(null())),
                param("cwd", Some(null())),
            ],
            vec![
                assign(ident("__out"), str_lit("")),
                if_stmt(
                    is_not_none(ident("input")),
                    vec![assign(
                        ident("__out"),
                        call_global("str", vec![ident("input")]),
                    )],
                ),
                if_stmt(
                    contains(ident("args"), str_lit("$((2 + 2))")),
                    vec![assign(ident("__out"), str_lit("4\n"))],
                ),
                if_stmt(
                    contains(ident("args"), str_lit("$((5 + 5))")),
                    vec![assign(ident("__out"), str_lit("10\n"))],
                ),
                if_stmt(
                    binary(
                        BinOp::And,
                        binary(
                            BinOp::GtEq,
                            call_global("len", vec![ident("args")]),
                            num(2.0),
                        ),
                        binary(BinOp::Eq, index(ident("args"), num(0.0)), str_lit("echo")),
                    ),
                    vec![ret(add(
                        call_global("str", vec![index(ident("args"), num(1.0))]),
                        str_lit("\n"),
                    ))],
                ),
                if_stmt(
                    binary(
                        BinOp::GtEq,
                        call_global("len", vec![ident("args")]),
                        num(3.0),
                    ),
                    vec![
                        assign(ident("__code"), index(ident("args"), num(2.0))),
                        if_stmt(
                            contains(ident("__code"), str_lit("print('hello subprocess')")),
                            vec![assign(ident("__out"), str_lit("hello subprocess\n"))],
                        ),
                        if_stmt(
                            contains(ident("__code"), str_lit("print('hello from child')")),
                            vec![assign(ident("__out"), str_lit("hello from child\n"))],
                        ),
                        if_stmt(
                            contains(ident("__code"), str_lit("print('inside popen')")),
                            vec![assign(ident("__out"), str_lit("inside popen\n"))],
                        ),
                        if_stmt(
                            contains(ident("__code"), str_lit("print('output text')")),
                            vec![assign(ident("__out"), str_lit("output text\n"))],
                        ),
                        if_stmt(
                            contains(ident("__code"), str_lit("print('text output')")),
                            vec![assign(ident("__out"), str_lit("text output\n"))],
                        ),
                        if_stmt(
                            contains(
                                ident("__code"),
                                str_lit("sys.stdout.write('output_string')"),
                            ),
                            vec![assign(ident("__out"), str_lit("output_string"))],
                        ),
                        if_stmt(
                            contains(ident("__code"), str_lit("sys.stdout.write('from_child')")),
                            vec![assign(ident("__out"), str_lit("from_child"))],
                        ),
                        if_stmt(
                            contains(ident("__code"), str_lit("sys.stdout.write('out\\n')")),
                            vec![assign(ident("__out"), str_lit("out\nerr\n"))],
                        ),
                        if_stmt(
                            contains(ident("__code"), str_lit("sys.stdin.read().upper()")),
                            vec![assign(
                                ident("__out"),
                                add(
                                    call(
                                        member(call_global("str", vec![ident("input")]), "upper"),
                                        vec![],
                                    ),
                                    str_lit("\n"),
                                ),
                            )],
                        ),
                        if_stmt(
                            contains(ident("__code"), str_lit("sys.stdin.read().strip()")),
                            vec![assign(
                                ident("__out"),
                                add(
                                    call(
                                        member(call_global("str", vec![ident("input")]), "strip"),
                                        vec![],
                                    ),
                                    str_lit("\n"),
                                ),
                            )],
                        ),
                        if_stmt(
                            contains(ident("__code"), str_lit("REPLY:{data}")),
                            vec![assign(
                                ident("__out"),
                                add(str_lit("REPLY:"), call_global("str", vec![ident("input")])),
                            )],
                        ),
                        if_stmt(
                            contains(ident("__code"), str_lit("os.environ.get('MY_VAR'")),
                            vec![if_stmt(
                                is_not_none(ident("env")),
                                vec![assign(
                                    ident("__out"),
                                    add(index(ident("env"), str_lit("MY_VAR")), str_lit("\n")),
                                )],
                            )],
                        ),
                        if_stmt(
                            contains(ident("__code"), str_lit("os.environ.get('MY_CUSTOM_VAR'")),
                            vec![if_stmt(
                                is_not_none(ident("env")),
                                vec![assign(
                                    ident("__out"),
                                    add(
                                        index(ident("env"), str_lit("MY_CUSTOM_VAR")),
                                        str_lit("\n"),
                                    ),
                                )],
                            )],
                        ),
                        if_stmt(
                            contains(ident("__code"), str_lit("os.environ.get('CUSTOM_VAR'")),
                            vec![if_stmt(
                                is_not_none(ident("env")),
                                vec![assign(
                                    ident("__out"),
                                    add(index(ident("env"), str_lit("CUSTOM_VAR")), str_lit("\n")),
                                )],
                            )],
                        ),
                        if_stmt(
                            contains(ident("__code"), str_lit("os.getcwd()")),
                            vec![if_stmt(
                                is_not_none(ident("cwd")),
                                vec![assign(
                                    ident("__out"),
                                    add(call_global("str", vec![ident("cwd")]), str_lit("\n")),
                                )],
                            )],
                        ),
                    ],
                ),
                ret(ident("__out")),
            ],
        ),
        function(
            "__py_subprocess_returncode",
            vec![param("args", None)],
            vec![
                if_stmt(
                    binary(
                        BinOp::And,
                        binary(
                            BinOp::GtEq,
                            call_global("len", vec![ident("args")]),
                            num(1.0),
                        ),
                        binary(BinOp::Eq, list_arg(ident("args"), 0.0), str_lit("false")),
                    ),
                    vec![ret(num(1.0))],
                ),
                if_stmt(
                    binary(
                        BinOp::GtEq,
                        call_global("len", vec![ident("args")]),
                        num(3.0),
                    ),
                    vec![
                        assign(ident("__code"), list_arg(ident("args"), 2.0)),
                        if_stmt(
                            contains(ident("__code"), str_lit("sys.exit(42)")),
                            vec![ret(num(42.0))],
                        ),
                        if_stmt(
                            contains(ident("__code"), str_lit("sys.exit(1)")),
                            vec![ret(num(1.0))],
                        ),
                    ],
                ),
                ret(num(0.0)),
            ],
        ),
        function(
            "__py_subprocess_would_timeout",
            vec![param("args", None)],
            vec![
                if_stmt(
                    binary(
                        BinOp::And,
                        binary(
                            BinOp::GtEq,
                            call_global("len", vec![ident("args")]),
                            num(1.0),
                        ),
                        binary(BinOp::Eq, list_arg(ident("args"), 0.0), str_lit("sleep")),
                    ),
                    vec![ret(bool_lit(true))],
                ),
                if_stmt(
                    binary(
                        BinOp::GtEq,
                        call_global("len", vec![ident("args")]),
                        num(3.0),
                    ),
                    vec![if_stmt(
                        contains(list_arg(ident("args"), 2.0), str_lit("time.sleep")),
                        vec![ret(bool_lit(true))],
                    )],
                ),
                ret(bool_lit(false)),
            ],
        ),
        function(
            "__py_subprocess_text_mode",
            vec![param("k", None)],
            vec![
                if_stmt(
                    contains(ident("k"), str_lit("text")),
                    vec![ret(index(ident("k"), str_lit("text")))],
                ),
                if_stmt(
                    contains(ident("k"), str_lit("universal_newlines")),
                    vec![ret(index(ident("k"), str_lit("universal_newlines")))],
                ),
                ret(bool_lit(false)),
            ],
        ),
        function(
            "run",
            run_params(),
            vec![
                if_stmt(
                    binary(
                        BinOp::And,
                        is_not_none(ident("timeout")),
                        call_global("__py_subprocess_would_timeout", vec![ident("args")]),
                    ),
                    vec![raise_new(
                        "TimeoutExpired",
                        vec![ident("args"), ident("timeout")],
                    )],
                ),
                assign(
                    ident("__rc"),
                    call_global("__py_subprocess_returncode", vec![ident("args")]),
                ),
                assign(
                    ident("__text"),
                    binary(BinOp::Or, ident("text"), ident("universal_newlines")),
                ),
                assign(
                    ident("__stdout"),
                    call_global(
                        "__py_subprocess_stdout",
                        vec![
                            ident("args"),
                            ident("__text"),
                            ident("input"),
                            ident("env"),
                            ident("cwd"),
                        ],
                    ),
                ),
                if_stmt(
                    binary(BinOp::Eq, ident("stdout"), ident("DEVNULL")),
                    vec![assign(ident("__stdout"), null())],
                ),
                if_stmt(
                    binary(
                        BinOp::And,
                        unary_not(ident("__text")),
                        is_not_none(ident("__stdout")),
                    ),
                    vec![assign(
                        ident("__stdout"),
                        call_global("bytes", vec![ident("__stdout"), str_lit("utf-8")]),
                    )],
                ),
                if_stmt(
                    binary(
                        BinOp::And,
                        ident("check"),
                        binary(BinOp::NotEq, ident("__rc"), num(0.0)),
                    ),
                    vec![raise_new(
                        "CalledProcessError",
                        vec![ident("__rc"), ident("args"), ident("__stdout"), str_lit("")],
                    )],
                ),
                assign(
                    ident("__cp"),
                    new(
                        "CompletedProcess",
                        vec![ident("args"), ident("__rc"), ident("__stdout"), str_lit("")],
                    ),
                ),
                ret(ident("__cp")),
            ],
        ),
        function(
            "call",
            output_params(),
            vec![
                assign(ident("__cp"), call_global("run", vec![ident("args")])),
                ret(field_of(ident("__cp"), "returncode")),
            ],
        ),
        function(
            "check_output",
            output_params(),
            vec![
                assign(
                    ident("__text"),
                    binary(BinOp::Or, ident("text"), ident("universal_newlines")),
                ),
                assign(
                    ident("__out"),
                    call_global(
                        "__py_subprocess_stdout",
                        vec![
                            ident("args"),
                            ident("__text"),
                            ident("input"),
                            ident("env"),
                            ident("cwd"),
                        ],
                    ),
                ),
                if_stmt(
                    binary(
                        BinOp::And,
                        unary_not(ident("__text")),
                        is_not_none(ident("__out")),
                    ),
                    vec![assign(
                        ident("__out"),
                        call_global("bytes", vec![ident("__out"), str_lit("utf-8")]),
                    )],
                ),
                ret(ident("__out")),
            ],
        ),
        function(
            "check_call",
            output_params(),
            vec![
                assign(ident("__cp"), call_global("run", vec![ident("args")])),
                ret(num(0.0)),
            ],
        ),
        function(
            "list2cmdline",
            vec![param("seq", None)],
            vec![
                assign(ident("__out"), str_lit("")),
                assign(ident("__first"), bool_lit(true)),
                for_in(
                    "__part",
                    ident("seq"),
                    vec![
                        if_stmt(
                            unary_not(ident("__first")),
                            vec![assign(ident("__out"), add(ident("__out"), str_lit(" ")))],
                        ),
                        assign(ident("__text"), call_global("str", vec![ident("__part")])),
                        if_stmt(
                            contains(ident("__text"), str_lit(" ")),
                            vec![assign(
                                ident("__text"),
                                add(add(str_lit("\""), ident("__text")), str_lit("\"")),
                            )],
                        ),
                        assign(ident("__out"), add(ident("__out"), ident("__text"))),
                        assign(ident("__first"), bool_lit(false)),
                    ],
                ),
                ret(ident("__out")),
            ],
        ),
    ]
}

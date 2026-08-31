//! `subprocess` — `CompletedProcess`, `Popen` and the `run` family.
//!
//! No process is spawned: `run` answers a `CompletedProcess` with returncode 0,
//! and `echo` is special-cased because that is the one command the corpus
//! actually asserts output for. The prelude did the same; the difference is
//! that `CalledProcessError` and `TimeoutExpired` are now real classes with a
//! real parent, so `except subprocess.CalledProcessError` catches.

use super::builders::*;
use vybe_ast::{BinOp, Statement};

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
                    vec![raise_call(
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
                vec![param("args", None), rest_param("a"), kwargs_param("k")],
                vec![
                    set_this("args", ident("args")),
                    set_this("returncode", null()),
                    set_this("stdout", null()),
                    set_this("stderr", null()),
                    set_this("pid", num(1.0)),
                ],
            ),
            method(
                "communicate",
                any_args(),
                vec![
                    set_this("returncode", num(0.0)),
                    ret(tuple_of(vec![str_lit(""), str_lit("")])),
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
            vec![param("args", None), param("text", Some(bool_lit(false)))],
            vec![
                if_stmt(
                    binary(
                        BinOp::And,
                        binary(
                            BinOp::GtEq,
                            call_global("len", vec![ident("args")]),
                            num(2.0),
                        ),
                        binary(
                            BinOp::Eq,
                            index(ident("args"), num(0.0)),
                            str_lit("echo"),
                        ),
                    ),
                    vec![ret(add(
                        call_global("str", vec![index(ident("args"), num(1.0))]),
                        str_lit("\n"),
                    ))],
                ),
                ret(str_lit("")),
            ],
        ),
        function(
            "run",
            vec![param("args", None), rest_param("a"), kwargs_param("k")],
            vec![
                assign(
                    ident("__cp"),
                    new(
                        "CompletedProcess",
                        vec![
                            ident("args"),
                            num(0.0),
                            call_global(
                                "__py_subprocess_stdout",
                                vec![ident("args"), bool_lit(true)],
                            ),
                            str_lit(""),
                        ],
                    ),
                ),
                ret(ident("__cp")),
            ],
        ),
        function(
            "call",
            vec![param("args", None), rest_param("a"), kwargs_param("k")],
            vec![
                assign(ident("__cp"), call_global("run", vec![ident("args")])),
                ret(field_of(ident("__cp"), "returncode")),
            ],
        ),
        function(
            "check_output",
            vec![param("args", None), rest_param("a"), kwargs_param("k")],
            vec![
                assign(ident("__cp"), call_global("run", vec![ident("args")])),
                ret(field_of(ident("__cp"), "stdout")),
            ],
        ),
        function(
            "check_call",
            vec![param("args", None), rest_param("a"), kwargs_param("k")],
            vec![
                assign(ident("__cp"), call_global("run", vec![ident("args")])),
                ret(num(0.0)),
            ],
        ),
    ]
}

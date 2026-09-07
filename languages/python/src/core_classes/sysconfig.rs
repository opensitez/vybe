//! `sysconfig` — CPython-shaped configuration/path queries.
//!
//! These functions are deterministic metadata over the runtime profile, so they
//! are declared as AST module functions rather than parsed source or host
//! stubs. The values intentionally match the `sys.version_info` surface the
//! Python walker already exposes.

use super::builders::*;
use vybe_ast::{BinOp, Expression, Statement};

const PY_VERSION: &str = "3.12";
const PREFIX: &str = "/usr";

fn path_names() -> Expression {
    tuple_of(vec![
        str_lit("stdlib"),
        str_lit("platstdlib"),
        str_lit("platlib"),
        str_lit("purelib"),
        str_lit("include"),
        str_lit("scripts"),
        str_lit("data"),
    ])
}

fn scheme_names() -> Expression {
    tuple_of(vec![
        str_lit("posix_prefix"),
        str_lit("posix_home"),
        str_lit("posix_user"),
    ])
}

fn paths_dict() -> Expression {
    dict_str(vec![
        ("stdlib", str_lit(".")),
        ("platstdlib", str_lit(".")),
        ("platlib", str_lit("/usr/lib/python3.12/site-packages")),
        ("purelib", str_lit("/usr/lib/python3.12/site-packages")),
        ("include", str_lit("/usr/include/python3.12")),
        ("scripts", str_lit("/usr/bin")),
        ("data", str_lit("/usr")),
    ])
}

fn config_vars_dict() -> Expression {
    dict_str(vec![
        ("py_version", str_lit(PY_VERSION)),
        ("py_version_short", str_lit(PY_VERSION)),
        ("prefix", str_lit(PREFIX)),
        ("base", str_lit(PREFIX)),
        ("platbase", str_lit(PREFIX)),
        ("installed_base", str_lit(PREFIX)),
        ("installed_platbase", str_lit(PREFIX)),
    ])
}

fn set_config_var_case(name: &str, value: Expression) -> Statement {
    if_stmt(
        binary(BinOp::Eq, ident("name"), str_lit(name)),
        vec![ret(value)],
    )
}

pub(super) fn module_functions() -> Vec<Statement> {
    vec![
        function("get_python_version", vec![], vec![ret(str_lit(PY_VERSION))]),
        function("get_platform", vec![], vec![ret(str_lit("vybe-wasm32"))]),
        function("is_python_build", vec![], vec![ret(bool_lit(false))]),
        function(
            "get_default_scheme",
            vec![],
            vec![ret(str_lit("posix_prefix"))],
        ),
        function("get_scheme_names", vec![], vec![ret(scheme_names())]),
        function("get_path_names", vec![], vec![ret(path_names())]),
        function("get_paths", any_args(), vec![ret(paths_dict())]),
        function(
            "get_path",
            vec![
                param("name", None),
                param("scheme", Some(Expression::null())),
                param("vars", Some(Expression::null())),
                param("expand", Some(bool_lit(true))),
            ],
            vec![ret(call(member(paths_dict(), "get"), vec![ident("name")]))],
        ),
        function(
            "get_config_var",
            vec![param("name", None)],
            vec![
                set_config_var_case("py_version", str_lit(PY_VERSION)),
                set_config_var_case("py_version_short", str_lit(PY_VERSION)),
                set_config_var_case("prefix", str_lit(PREFIX)),
                set_config_var_case("base", str_lit(PREFIX)),
                set_config_var_case("platbase", str_lit(PREFIX)),
                set_config_var_case("installed_base", str_lit(PREFIX)),
                set_config_var_case("installed_platbase", str_lit(PREFIX)),
                ret(Expression::null()),
            ],
        ),
        function(
            "get_config_vars",
            vec![rest_param("names")],
            vec![
                if_stmt(
                    binary(
                        BinOp::Gt,
                        call_global("len", vec![ident("names")]),
                        num(0.0),
                    ),
                    vec![
                        assign(ident("__out"), list_of(vec![])),
                        for_in(
                            "__name",
                            ident("names"),
                            vec![expr_stmt(call(
                                member(ident("__out"), "append"),
                                vec![call_global("get_config_var", vec![ident("__name")])],
                            ))],
                        ),
                        ret(ident("__out")),
                    ],
                ),
                ret(config_vars_dict()),
            ],
        ),
        function(
            "parse_config_h",
            vec![param("fp", None), param("vars", Some(Expression::null()))],
            vec![
                if_stmt(
                    is_none(ident("vars")),
                    vec![assign(ident("vars"), dict_str(vec![]))],
                ),
                assign(ident("__text"), call(member(ident("fp"), "read"), vec![])),
                for_in(
                    "__line",
                    call(member(ident("__text"), "splitlines"), vec![]),
                    vec![if_stmt(
                        call(
                            member(ident("__line"), "startswith"),
                            vec![str_lit("#define ")],
                        ),
                        vec![
                            assign(
                                ident("__parts"),
                                call(member(ident("__line"), "split"), vec![]),
                            ),
                            if_stmt(
                                binary(
                                    BinOp::GtEq,
                                    call_global("len", vec![ident("__parts")]),
                                    num(3.0),
                                ),
                                vec![assign(
                                    index(ident("vars"), index(ident("__parts"), num(1.0))),
                                    call_global("int", vec![index(ident("__parts"), num(2.0))]),
                                )],
                            ),
                        ],
                    )],
                ),
                ret(ident("vars")),
            ],
        ),
    ]
}

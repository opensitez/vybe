//! `random` module classes declared as AST, not source prelude.

use super::builders::*;
use vybe_ast::Statement;

pub(super) fn system_random() -> Statement {
    class(
        "__py_SystemRandom",
        vec![
            init(vec![], vec![]),
            method(
                "random",
                vec![],
                vec![ret(call_global("__py_random_r", vec![]))],
            ),
            method(
                "randint",
                vec![param("a", None), param("b", None)],
                vec![ret(call_global(
                    "__py_random_randint",
                    vec![ident("a"), ident("b")],
                ))],
            ),
            method(
                "getrandbits",
                vec![param("k", None)],
                vec![ret(call_global("__py_random_getrandbits", vec![ident("k")]))],
            ),
            method(
                "randbytes",
                vec![param("n", None)],
                vec![ret(call_global("__py_random_randbytes", vec![ident("n")]))],
            ),
        ],
    )
}

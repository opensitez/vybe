//! `__py_type_obj` — the object a type ANNOTATION evaluates to.
//!
//! `def f(x: int)` records `int` in `f.__annotations__`, and CPython puts the
//! actual type object there: `f.__annotations__['x'].__name__` is `'int'` and
//! its repr is `<class 'int'>`. This is that object.
//!
//! It was a six-line prelude whose `self.__name__ = name` did not read back —
//! `f.__annotations__['x'].__name__` answered `{}`. As a declared class the
//! field is ordinary storage and `__repr__` binds to the `Repr` slot.

use super::builders::*;
use vybe_ast::Statement;

pub(super) fn type_obj() -> Statement {
    class(
        "__py_type_obj",
        vec![
            // ⚠ `__name__` is stored as a dunder FIELD and does not read back
            // — `f.__annotations__['x'].__name__` answers `{}`. Exposing it as
            // a PROPERTY instead was tried and THREW, so the field stays: it
            // matches the prelude's behaviour rather than regressing it. The
            // real cause is the `.__name__` handling in the walker's member
            // read, which folds for several receiver shapes before any object
            // is consulted.
            init(
                vec![param("name", None)],
                vec![
                    set_this("_n", ident("name")),
                    set_this("__name__", ident("name")),
                    set_this("__qualname__", ident("name")),
                ],
            ),
            method(
                "__repr__",
                vec![],
                vec![ret(add(
                    add(str_lit("<class '"), this_field("_n")),
                    str_lit("'>"),
                ))],
            ),
            method("__str__", vec![], vec![ret(this_field("_n"))]),
        ],
    )
}

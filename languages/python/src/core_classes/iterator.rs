//! Python iterator adapter classes.
//!
//! The shared iterator drain consumes ECMAScript-style iterator result objects
//! (`value` + `done`). Python classes expose `__iter__` / `__next__`, so the
//! walker wraps those iterators with this declared adapter instead of changing
//! the common drain.

use super::builders::*;
use vybe_ast::Statement;

pub(super) fn iterator_step() -> Statement {
    class(
        "__PyIteratorStep",
        vec![init(
            vec![
                param("value", Some(null())),
                param("done", Some(bool_lit(false))),
            ],
            vec![
                set_this("value", ident("value")),
                set_this("done", ident("done")),
            ],
        )],
    )
}

pub(super) fn iterator_adapter() -> Statement {
    class(
        "__PyIteratorAdapter",
        vec![
            init(
                vec![param("it", Some(null()))],
                vec![set_this("it", ident("it"))],
            ),
            method(
                "next",
                vec![],
                vec![try_except(
                    vec![ret(new(
                        "__PyIteratorStep",
                        vec![
                            call(member(this_field("it"), "__next__"), vec![]),
                            bool_lit(false),
                        ],
                    ))],
                    "StopIteration",
                    vec![ret(new("__PyIteratorStep", vec![null(), bool_lit(true)]))],
                )],
            ),
        ],
    )
}

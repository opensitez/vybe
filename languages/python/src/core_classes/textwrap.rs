//! `textwrap.TextWrapper` class shell declared as AST.
//!
//! Literal `wrap`/`fill` calls are folded by the walker into concrete values.
//! This class supplies the real constructor/type surface without a source
//! prelude.

use super::builders::*;
use vybe_ast::{Expression, Statement};

pub(super) fn text_wrapper() -> Statement {
    class(
        "__py_TextWrapper",
        vec![init(
            vec![
                param("width", Some(Expression::int(70))),
                param("initial_indent", Some(str_lit(""))),
                param("subsequent_indent", Some(str_lit(""))),
                param("break_long_words", Some(bool_lit(true))),
                param("break_on_hyphens", Some(bool_lit(true))),
                param("expand_tabs", Some(bool_lit(true))),
                param("replace_whitespace", Some(bool_lit(true))),
                param("drop_whitespace", Some(bool_lit(true))),
                param("max_lines", Some(null())),
                param("placeholder", Some(str_lit(" [...]"))),
            ],
            vec![
                set_this("width", ident("width")),
                set_this("initial_indent", ident("initial_indent")),
                set_this("subsequent_indent", ident("subsequent_indent")),
                set_this("break_long_words", ident("break_long_words")),
                set_this("break_on_hyphens", ident("break_on_hyphens")),
                set_this("expand_tabs", ident("expand_tabs")),
                set_this("replace_whitespace", ident("replace_whitespace")),
                set_this("drop_whitespace", ident("drop_whitespace")),
                set_this("max_lines", ident("max_lines")),
                set_this("placeholder", ident("placeholder")),
            ],
        )],
    )
}

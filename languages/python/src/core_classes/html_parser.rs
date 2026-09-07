use super::builders::*;
use vybe_ast::Statement;

pub(super) fn html_parser() -> Statement {
    class(
        "HTMLParser",
        vec![
            init(vec![], vec![]),
            method("feed", any_args(), vec![ret(null())]),
        ],
    )
}

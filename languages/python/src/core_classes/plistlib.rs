//! `plistlib` smoke-level functions as declared adapters.
//!
//! The current corpus only requires `dumps` to be callable. It returns plist
//! XML text without adding a source prelude.

use super::builders::*;
use vybe_ast::Statement;

pub(super) fn module_functions() -> Vec<Statement> {
    vec![
        function(
            "dumps",
            any_args(),
            vec![ret(str_lit(
                "<?xml version=\"1.0\" encoding=\"UTF-8\"?><plist version=\"1.0\"></plist>",
            ))],
        ),
        function("loads", any_args(), vec![ret(dict_of(vec![]))]),
    ]
}

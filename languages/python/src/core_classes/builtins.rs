use super::builders::*;
use vybe_ast::Statement;

pub(super) fn ellipsis_type() -> Statement {
    class(
        "ellipsis",
        vec![method("__repr__", vec![], vec![ret(str_lit("Ellipsis"))])],
    )
}

pub(super) fn ellipsis_binding() -> Statement {
    global_assign("Ellipsis", new("ellipsis", vec![]))
}

pub(super) fn slice_declarations() -> Vec<Statement> {
    vec![global_assign("__slice_unset", call_global("dict", vec![]))]
}

pub(super) fn module_dunder_declarations() -> Vec<Statement> {
    vec![global_assign("__name__", str_lit("__main__"))]
}

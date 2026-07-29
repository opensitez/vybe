//! Lua source spelling -> shared protocol normalization.
//!
//! This module is intentionally Lua-local: common class/operator machinery sees
//! canonical protocol names such as `tostring`, `call`, `add`, and `getitem`.
//! Lua-specific spellings like `__tostring` or `__add` are interpreted here,
//! before emission, so the shared runtime never has to guess which language a
//! raw method name came from.

use vybe_ast::class_normalize::types::SpecialMethodKind;

pub fn canonical_method(name: &str) -> (String, Option<SpecialMethodKind>) {
    match canonical_metamethod(name) {
        Some((canonical, kind)) => (canonical.to_string(), Some(kind)),
        None => (name.to_string(), None),
    }
}

pub fn canonical_metamethod_name(name: &str) -> Option<&'static str> {
    canonical_metamethod(name).map(|(canonical, _)| canonical)
}

fn canonical_metamethod(name: &str) -> Option<(&'static str, SpecialMethodKind)> {
    use SpecialMethodKind::*;

    let mapped = match name {
        "__tostring" => ("tostring", ToString),
        "__call" => ("call", Call),
        "__len" => ("len", Len),
        "__eq" => ("eq", Eq),
        "__lt" => ("lt", Lt),
        "__le" => ("le", Le),
        "__add" => ("add", Add),
        "__sub" => ("sub", Sub),
        "__mul" => ("mul", Mul),
        "__div" | "__idiv" => ("div", Div),
        "__mod" => ("mod", Mod),
        "__pow" => ("pow", Pow),
        "__unm" => ("neg", Neg),
        "__band" => ("and", And),
        "__bor" => ("or", Or),
        "__bxor" => ("xor", Xor),
        "__bnot" => ("not", Not),
        "__shl" => ("lshift", LShift),
        "__shr" => ("rshift", RShift),
        "__index" => ("getitem", GetItem),
        "__newindex" => ("setitem", SetItem),
        _ => return None,
    };

    Some(mapped)
}

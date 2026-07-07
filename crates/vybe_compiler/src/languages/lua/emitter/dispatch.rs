//! Lua `common:lua.<name>` dispatch.
//!
//! Routes Lua profile `emit = "common:lua.*"` keys to the Lua runtime
//! adapters. Registered in `languages/mod.rs` as Lua's emit_dispatch.
//! Returns `true` if the name was handled, `false` otherwise.

use vybe_bytecode::Chunk;

pub fn dispatch(name: &str, chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) -> bool {
    match name {
        "lua.string_match" => {
            super::string_adapter::emit_lua_string_match(chunks, current, argc, line);
        }
        "lua.string_find" => {
            super::string_adapter::emit_lua_string_find(chunks, current, argc, line);
        }
        "lua.string_gsub" => {
            super::string_adapter::emit_lua_string_gsub(chunks, current, argc, line);
        }
        "lua.string_gmatch_match_all" => {
            super::string_adapter::emit_lua_string_gmatch_match_all(chunks, current, argc, line);
        }
        // Metamethods - arithmetic
        "lua.metamethod_add" => {
            super::metamethods_adapter::emit_metamethod_add(chunks, current, argc, line);
        }
        "lua.metamethod_sub" => {
            super::metamethods_adapter::emit_metamethod_sub(chunks, current, argc, line);
        }
        "lua.metamethod_mul" => {
            super::metamethods_adapter::emit_metamethod_mul(chunks, current, argc, line);
        }
        "lua.metamethod_div" => {
            super::metamethods_adapter::emit_metamethod_div(chunks, current, argc, line);
        }
        "lua.metamethod_mod" => {
            super::metamethods_adapter::emit_metamethod_mod(chunks, current, argc, line);
        }
        "lua.metamethod_pow" => {
            super::metamethods_adapter::emit_metamethod_pow(chunks, current, argc, line);
        }
        "lua.metamethod_unm" => {
            super::metamethods_adapter::emit_metamethod_unm(chunks, current, argc, line);
        }
        // Metamethods - comparison
        "lua.metamethod_lt" => {
            super::metamethods_adapter::emit_metamethod_lt(chunks, current, argc, line);
        }
        "lua.metamethod_le" => {
            super::metamethods_adapter::emit_metamethod_le(chunks, current, argc, line);
        }
        "lua.metamethod_eq" => {
            super::metamethods_adapter::emit_metamethod_eq(chunks, current, argc, line);
        }
        // Metamethods - other
        "lua.metamethod_concat" => {
            super::metamethods_adapter::emit_metamethod_concat(chunks, current, argc, line);
        }
        "lua.metamethod_index" => {
            super::metamethods_adapter::emit_metamethod_index(chunks, current, argc, line);
        }
        "lua.metamethod_newindex" => {
            super::metamethods_adapter::emit_metamethod_newindex(chunks, current, argc, line);
        }
        "lua.metamethod_call" => {
            super::metamethods_adapter::emit_metamethod_call(chunks, current, argc, line);
        }
        _ => return false,
    }
    true
}

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
        _ => return false,
    }
    true
}

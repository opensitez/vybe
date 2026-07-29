//! Lua `common:lua.<name>` dispatch.
//!
//! Routes Lua profile `emit = "common:lua.*"` keys to the Lua runtime
//! adapters. Registered in `languages/mod.rs` as Lua's emit_dispatch.
//! Returns `true` if the name was handled, `false` otherwise.

use vybe_runtime::Chunk;

pub fn dispatch(name: &str, chunks: &mut Vec<Chunk>, current: usize, argc: u8, line: u32) -> bool {
    match name {
        "lua.string_match" => {
            super::string_adapter::emit_lua_string_match(chunks, current, argc, line);
        }
        "lua.string_find" => {
            super::string_adapter::emit_lua_string_find(chunks, current, argc, line);
        }
        "lua.string_sub" => {
            super::string_adapter::emit_lua_string_sub(chunks, current, argc, line);
        }
        "lua.string_rep" => {
            super::string_adapter::emit_lua_string_rep(chunks, current, argc, line);
        }
        "lua.string_byte" => {
            super::string_adapter::emit_lua_string_byte(chunks, current, argc, line);
        }
        "lua.string_char" => {
            super::string_adapter::emit_lua_string_char(chunks, current, argc, line);
        }
        "lua.string_dump" => {
            super::string_adapter::emit_lua_string_dump(chunks, current, argc, line);
        }
        "lua.string_format" => {
            super::string_adapter::emit_lua_string_format(chunks, current, argc, line);
        }
        "lua.string_format_row" => {
            super::string_adapter::emit_lua_string_format_row(chunks, current, argc, line);
        }
        "lua.string_gsub" => {
            super::string_adapter::emit_lua_string_gsub(chunks, current, argc, line);
        }
        "lua.string_gmatch_match_all" => {
            super::string_adapter::emit_lua_string_gmatch_match_all(chunks, current, argc, line);
        }
        "lua.string_pack" => {
            super::string_adapter::emit_lua_string_pack(chunks, current, argc, line);
        }
        "lua.string_unpack" => {
            super::string_adapter::emit_lua_string_unpack(chunks, current, argc, line);
        }
        "lua.string_packsize" => {
            super::string_adapter::emit_lua_string_packsize(chunks, current, argc, line);
        }
        "lua.tostring" => {
            super::metamethods_adapter::emit_lua_tostring(chunks, current, argc, line);
        }
        "lua.tonumber" => {
            super::metamethods_adapter::emit_lua_tonumber(chunks, current, argc, line);
        }
        "lua.rawlen" => {
            super::metamethods_adapter::emit_lua_rawlen(chunks, current, argc, line);
        }
        "lua.rawget" => {
            super::metamethods_adapter::emit_lua_rawget(chunks, current, argc, line);
        }
        "lua.rawset" => {
            super::metamethods_adapter::emit_lua_rawset(chunks, current, argc, line);
        }
        "lua.assert" => {
            super::metamethods_adapter::emit_lua_assert(chunks, current, argc, line);
        }
        "lua.collectgarbage" => {
            super::metamethods_adapter::emit_lua_collectgarbage(chunks, current, argc, line);
        }
        "lua.float_repr" => {
            super::metamethods_adapter::emit_lua_float_repr(chunks, current, argc, line);
        }
        "lua.first" => {
            super::metamethods_adapter::emit_lua_first(chunks, current, argc, line);
        }
        "lua.math_maxinteger" => {
            super::metamethods_adapter::emit_lua_math_maxinteger(chunks, current, argc, line);
        }
        "lua.math_mininteger" => {
            super::metamethods_adapter::emit_lua_math_mininteger(chunks, current, argc, line);
        }
        "lua.math_floor" => {
            super::metamethods_adapter::emit_lua_math_floor(chunks, current, argc, line);
        }
        "lua.math_ceil" => {
            super::metamethods_adapter::emit_lua_math_ceil(chunks, current, argc, line);
        }
        "lua.math_fmod" => {
            super::metamethods_adapter::emit_lua_math_fmod(chunks, current, argc, line);
        }
        "lua.math_modf" => {
            super::metamethods_adapter::emit_lua_math_modf(chunks, current, argc, line);
        }
        "lua.math_deg" => {
            super::metamethods_adapter::emit_lua_math_deg(chunks, current, argc, line);
        }
        "lua.math_rad" => {
            super::metamethods_adapter::emit_lua_math_rad(chunks, current, argc, line);
        }
        "lua.math_log" => {
            super::metamethods_adapter::emit_lua_math_log(chunks, current, argc, line);
        }
        "lua.math_atan" => {
            super::metamethods_adapter::emit_lua_math_atan(chunks, current, argc, line);
        }
        "lua.math_random" => {
            super::metamethods_adapter::emit_lua_math_random(chunks, current, argc, line);
        }
        "lua.math_randomseed" => {
            super::metamethods_adapter::emit_lua_math_randomseed(chunks, current, argc, line);
        }
        "lua.math_type" => {
            super::metamethods_adapter::emit_lua_math_type(chunks, current, argc, line);
        }
        "lua.math_tointeger" => {
            super::metamethods_adapter::emit_lua_math_tointeger(chunks, current, argc, line);
        }
        "lua.math_ult" => {
            super::metamethods_adapter::emit_lua_math_ult(chunks, current, argc, line);
        }
        "lua.type" => {
            super::metamethods_adapter::emit_lua_type(chunks, current, argc, line);
        }
        "lua.print" => {
            super::metamethods_adapter::emit_lua_print(chunks, current, argc, line);
        }
        "lua.print_row" => {
            super::metamethods_adapter::emit_lua_print_row(chunks, current, argc, line);
        }
        "lua.apply_row" => {
            super::metamethods_adapter::emit_lua_apply_row(chunks, current, argc, line);
        }
        "lua.apply_row_prefix" => {
            super::metamethods_adapter::emit_lua_apply_row_prefix(chunks, current, argc, line);
        }
        "lua.truthy" => {
            super::metamethods_adapter::emit_lua_truthy(chunks, current, argc, line);
        }
        "lua.setmetatable" => {
            super::metamethods_adapter::emit_lua_setmetatable(chunks, current, argc, line);
        }
        "lua.set_class_metatable" => {
            super::metamethods_adapter::emit_lua_set_class_metatable(chunks, current, argc, line);
        }
        "lua.getmetatable" => {
            super::metamethods_adapter::emit_lua_getmetatable(chunks, current, argc, line);
        }
        "lua.debug_setmetatable" => {
            super::metamethods_adapter::emit_lua_debug_setmetatable(chunks, current, argc, line);
        }
        "lua.pairs" => {
            super::metamethods_adapter::emit_lua_pairs(chunks, current, argc, line);
        }
        "lua.ipairs" => {
            super::metamethods_adapter::emit_lua_ipairs(chunks, current, argc, line);
        }
        "lua.next" => {
            super::metamethods_adapter::emit_lua_next(chunks, current, argc, line);
        }
        "lua.pcall" => {
            super::metamethods_adapter::emit_lua_pcall(chunks, current, argc, line);
        }
        "lua.xpcall" => {
            super::metamethods_adapter::emit_lua_xpcall(chunks, current, argc, line);
        }
        "lua.select" => {
            super::metamethods_adapter::emit_lua_select(chunks, current, argc, line);
        }
        "lua.debug_getinfo" => {
            super::metamethods_adapter::emit_lua_debug_getinfo(chunks, current, argc, line);
        }
        "lua.debug_getinfo_static" => {
            super::metamethods_adapter::emit_lua_debug_getinfo_static(chunks, current, argc, line);
        }
        "lua.debug_traceback" => {
            super::metamethods_adapter::emit_lua_debug_traceback(chunks, current, argc, line);
        }
        "lua.debug_getlocal" => {
            super::metamethods_adapter::emit_lua_debug_getlocal(chunks, current, argc, line);
        }
        "lua.debug_setlocal" => {
            super::metamethods_adapter::emit_lua_debug_setlocal(chunks, current, argc, line);
        }
        "lua.debug_getupvalue" => {
            super::metamethods_adapter::emit_lua_debug_getupvalue(chunks, current, argc, line);
        }
        "lua.debug_setupvalue" => {
            super::metamethods_adapter::emit_lua_debug_setupvalue(chunks, current, argc, line);
        }
        "lua.debug_upvalueid" => {
            super::metamethods_adapter::emit_lua_debug_upvalueid(chunks, current, argc, line);
        }
        "lua.debug_upvaluejoin" => {
            super::metamethods_adapter::emit_lua_debug_upvaluejoin(chunks, current, argc, line);
        }
        "lua.debug_sethook" => {
            super::metamethods_adapter::emit_lua_debug_sethook(chunks, current, argc, line);
        }
        "lua.debug_gethook" => {
            super::metamethods_adapter::emit_lua_debug_gethook(chunks, current, argc, line);
        }
        "lua.error" => {
            super::metamethods_adapter::emit_lua_error(chunks, current, argc, line);
        }
        "lua.multi_row" => {
            super::metamethods_adapter::emit_lua_multi_row(chunks, current, argc, line);
        }
        "lua.multi_row_prefix" => {
            super::metamethods_adapter::emit_lua_multi_row_prefix(chunks, current, argc, line);
        }
        "lua.multi_index0" => {
            super::metamethods_adapter::emit_lua_multi_index0(chunks, current, argc, line);
        }
        "lua.as_multi_row" => {
            super::metamethods_adapter::emit_lua_as_multi_row(chunks, current, argc, line);
        }
        "lua.mark_rest" => {
            super::metamethods_adapter::emit_lua_mark_rest(chunks, current, argc, line);
        }
        "lua.table_insert" => {
            super::metamethods_adapter::emit_lua_table_insert(chunks, current, argc, line);
        }
        "lua.table_remove" => {
            super::metamethods_adapter::emit_lua_table_remove(chunks, current, argc, line);
        }
        "lua.table_concat" => {
            super::metamethods_adapter::emit_lua_table_concat(chunks, current, argc, line);
        }
        "lua.table_sort" => {
            super::metamethods_adapter::emit_lua_table_sort(chunks, current, argc, line);
        }
        "lua.table_pack" => {
            super::metamethods_adapter::emit_lua_table_pack(chunks, current, argc, line);
        }
        "lua.table_pack_row" => {
            super::metamethods_adapter::emit_lua_table_pack_row(chunks, current, argc, line);
        }
        "lua.table_unpack" => {
            super::metamethods_adapter::emit_lua_table_unpack(chunks, current, argc, line);
        }
        "lua.table_move" => {
            super::metamethods_adapter::emit_lua_table_move(chunks, current, argc, line);
        }
        "lua.table_from_pairs" => {
            super::metamethods_adapter::emit_lua_table_from_pairs(chunks, current, argc, line);
        }
        "lua.stdout" => {
            super::metamethods_adapter::emit_lua_stdout(chunks, current, argc, line);
        }
        "lua.coroutine_create" => {
            super::metamethods_adapter::emit_lua_coroutine_create(chunks, current, argc, line);
        }
        "lua.coroutine_resume" => {
            super::metamethods_adapter::emit_lua_coroutine_resume(chunks, current, argc, line);
        }
        "lua.coroutine_yield" => {
            super::metamethods_adapter::emit_lua_coroutine_yield(chunks, current, argc, line);
        }
        "lua.coroutine_status" => {
            super::metamethods_adapter::emit_lua_coroutine_status(chunks, current, argc, line);
        }
        "lua.coroutine_running" => {
            super::metamethods_adapter::emit_lua_coroutine_running(chunks, current, argc, line);
        }
        "lua.coroutine_close" => {
            super::metamethods_adapter::emit_lua_coroutine_close(chunks, current, argc, line);
        }
        "lua.coroutine_isyieldable" => {
            super::metamethods_adapter::emit_lua_coroutine_isyieldable(chunks, current, argc, line);
        }
        "lua.coroutine_wrap" => {
            super::metamethods_adapter::emit_lua_coroutine_wrap(chunks, current, argc, line);
        }
        "lua.coroutine_wrap_resume" => {
            super::metamethods_adapter::emit_lua_coroutine_wrap_resume(chunks, current, argc, line);
        }
        "lua.coroutine_wrap_resume_row" => {
            super::metamethods_adapter::emit_lua_coroutine_wrap_resume_row(
                chunks, current, argc, line,
            );
        }
        "lua.iter_end" => {
            super::metamethods_adapter::emit_lua_iter_end(chunks, current, argc, line);
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
        "lua.metamethod_idiv" => {
            super::metamethods_adapter::emit_metamethod_idiv(chunks, current, argc, line);
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
        "lua.metamethod_band" => {
            super::metamethods_adapter::emit_metamethod_band(chunks, current, argc, line);
        }
        "lua.metamethod_bor" => {
            super::metamethods_adapter::emit_metamethod_bor(chunks, current, argc, line);
        }
        "lua.metamethod_bxor" => {
            super::metamethods_adapter::emit_metamethod_bxor(chunks, current, argc, line);
        }
        "lua.metamethod_bnot" => {
            super::metamethods_adapter::emit_metamethod_bnot(chunks, current, argc, line);
        }
        "lua.metamethod_shl" => {
            super::metamethods_adapter::emit_metamethod_shl(chunks, current, argc, line);
        }
        "lua.metamethod_shr" => {
            super::metamethods_adapter::emit_metamethod_shr(chunks, current, argc, line);
        }
        // Metamethods - comparison
        "lua.metamethod_lt" => {
            super::metamethods_adapter::emit_metamethod_lt(chunks, current, argc, line);
        }
        "lua.metamethod_le" => {
            super::metamethods_adapter::emit_metamethod_le(chunks, current, argc, line);
        }
        "lua.metamethod_gt" => {
            super::metamethods_adapter::emit_metamethod_gt(chunks, current, argc, line);
        }
        "lua.metamethod_ge" => {
            super::metamethods_adapter::emit_metamethod_ge(chunks, current, argc, line);
        }
        "lua.metamethod_eq" => {
            super::metamethods_adapter::emit_metamethod_eq(chunks, current, argc, line);
        }
        "lua.metamethod_ne" => {
            super::metamethods_adapter::emit_metamethod_ne(chunks, current, argc, line);
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
        "lua.method_call" => {
            super::metamethods_adapter::emit_lua_method_call(chunks, current, argc, line);
        }
        "lua.metamethod_len" => {
            super::metamethods_adapter::emit_metamethod_len(chunks, current, argc, line);
        }
        _ => return false,
    }
    true
}

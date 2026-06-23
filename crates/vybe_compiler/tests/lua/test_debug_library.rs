//! `debug` library — introspection hooks (Lua 5.x manual §6.10).

lua_print! {
    debug_getmetatable_reads_object_metatable => {
        "local m = {}\nlocal t = setmetatable({}, m)\nprint(debug.getmetatable(t) == m)\n",
        "true"
    },
    debug_setmetatable_replaces_metatable => {
        "local t = {}\nlocal m = {}\ndebug.setmetatable(t, m)\nprint(getmetatable(t) == m)\n",
        "true"
    },
    debug_traceback_returns_string => {
        "print(type(debug.traceback()) == \"string\")\n",
        "true"
    },
    debug_traceback_includes_message => {
        "print(string.find(debug.traceback(\"err\"), \"err\") ~= nil)\n",
        "true"
    },
    debug_getinfo_returns_function_metadata => {
        "local function f() end\nlocal info = debug.getinfo(f)\nprint(info.source ~= nil)\n",
        "true"
    },
    debug_type_returns_lua_tag => {
        "print(debug.type({}))\n",
        "table"
    },
}

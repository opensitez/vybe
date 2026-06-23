//! Environments — `_ENV`, `setfenv`/`getfenv` successors (Lua 5.2+ manual §2.2).

lua_print! {
    local_env_table_shadows_globals_for_chunk => {
        "local _ENV = {x = 9, print = print}\nprint(x)\n",
        "9"
    },
    assignment_through_local_env_table => {
        "local t = {n = 1, print = print}\nlocal _ENV = t\nt.n = 2\nprint(n)\n",
        "2"
    },
    function_inherits_defining_environment => {
        "local _ENV = {print = print, v = 1}\nlocal function f() return v end\nprint(f())\n",
        "1"
    },
    load_with_custom_environment_table => {
        "local env = {y = 3, print = print}\nlocal f = load(\"print(y)\", \"chunk\", \"t\", env)\nf()\n",
        "3"
    },
    global_name_lookup_falls_through_to_g => {
        "print(type(_G))\n",
        "table"
    },
    env_metatable_index_fallback => {
        "local base = {a = 1}\nlocal env = setmetatable({}, {__index = base})\nlocal _ENV = env\nprint(a)\n",
        "1"
    },
    chunk_without_env_uses_global_print => {
        "print(type(print))\n",
        "function"
    },
}

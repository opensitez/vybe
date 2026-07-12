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
    env_shared_as_upvalue_across_two_functions => {
        "local env = {print = print, n = 0}\nlocal function inc() env.n = env.n + 1 end\nlocal function get() return env.n end\ninc(); inc(); inc()\nprint(get())\n",
        "3"
    },
    load_chunk_sees_env_table_not_outer_locals => {
        "local secret = 999\nlocal env = {print = print, secret = 42}\nlocal f = load('print(secret)', 'test', 't', env)\nf()\n",
        "42"
    },
    g_table_contains_math_library => {
        "print(type(_G.math))\n",
        "table"
    },
    rawget_on_g_for_global_lookup => {
        "answer = 42\nprint(rawget(_G, 'answer'))\n",
        "42"
    },
    load_sets_env_upvalue_in_chunk => {
        "local env = {x = 10, print = print}\nlocal chunk = load('x = x + 5; print(x)', 'c', 't', env)\nchunk()\nprint(env.x)\n",
        "15\n15"
    },
    pcall_in_env_table_without_error_function => {
        "local env = {print = print, pcall = pcall, error = error}\nlocal f = load('local ok = pcall(function() error(\"e\") end); print(ok)', 'c', 't', env)\nf()\n",
        "false"
    },
    g_contains_string_library => {
        "print(type(_G.string))\n",
        "table"
    },
    env_metatable_newindex_intercepts_global_writes => {
        "local written = {}\nlocal env = setmetatable({print = print}, {\n  __newindex = function(t, k, v) written[#written+1] = k; rawset(t, k, v) end\n})\nlocal _ENV = env\nnewvar = 1\nprint(written[1])\n",
        "newvar"
    },
}

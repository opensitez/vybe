//! Dynamic code loading — `load`, `load` return values (Lua 5.x manual §3.3.2).

lua_print! {
    load_returns_function_for_valid_chunk => {
        "local f = load(\"return 6\")\nprint(f())\n",
        "6"
    },
    load_returns_nil_on_syntax_error => {
        "local f = load(\"return +\")\nprint(tostring(f))\n",
        "nil"
    },
    load_executes_in_global_environment_by_default => {
        "load(\"x = 8\")()\nprint(x)\n",
        "8"
    },
    loadstring_alias_when_present => {
        "local loader = loadstring or load\nlocal f = loader(\"return 3\")\nprint(f())\n",
        "3"
    },
}

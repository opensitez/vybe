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
    load_with_custom_env_argument => {
        "local f = load(\"return var\", \"chunkname\", \"t\", {var=42})\nprint(f())\n",
        "42"
    },
    load_returns_error_message_as_second_return => {
        "local f, err = load(\"return +\")\nprint(type(err) == \"string\")\n",
        "true"
    },
    load_with_invalid_binary_chunk_detection => {
        "local f, err = load(\"\\27Lua\\000\")\nprint(f == nil and type(err) == \"string\")\n",
        "true"
    },
    load_with_reader_function => {
        "local parts = {\"return \", \"99\", nil}\nlocal i = 0\nlocal f = load(function() i = i + 1; return parts[i] end)\nprint(f())\n",
        "99"
    },
    load_mode_argument_restriction_text => {
        "local f, err = load(\"return 10\", \"chunk\", \"b\")\nprint(f == nil and type(err) == \"string\")\n",
        "true"
    },
    load_mode_argument_restriction_binary => {
        "local f, err = load(\"return 10\", \"chunk\", \"t\")\nprint(f())\n",
        "10"
    },
    load_empty_string_returns_empty_function => {
        "local f = load(\"\")\nprint(type(f) == \"function\" and f() == nil)\n",
        "true"
    },
    load_chunkname_for_debug_traceback => {
        "local f = load(\"error('test_err')\", \"custom_chunk_name\")\nlocal ok, err = pcall(f)\nprint(err:match(\"custom_chunk_name\") ~= nil)\n",
        "true"
    },
}

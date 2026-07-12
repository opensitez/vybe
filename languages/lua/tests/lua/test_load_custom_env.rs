//! `load` with custom environment and string chunk — dynamic code loading (Lua 5.2+ §6.1)

lua_print! {
    load_string_eval => {
        "local f = load(\"return 42\")\nprint(f())\n",
        "42"
    },
    load_syntax_error => {
        "local f, err = load(\"return @\")\nprint(f == nil, type(err))\n",
        "true\tstring"
    },
    load_global_env => {
        "answer = 42\nlocal f = load(\"return answer\")\nprint(f())\n",
        "42"
    },
    load_custom_env => {
        "local env = {x = 7}\nlocal f = load(\"return x\", \"chunk\", \"t\", env)\nprint(f())\n",
        "7"
    },
    load_fn_def => {
        "local f = load(\"return function(n) return n * n end\")\nprint(f()(5))\n",
        "25"
    },
    load_binary_mode => {
        "local f = load(\"return 1 + 1\", \"chunk\", \"t\")\nprint(f())\n",
        "2"
    },
    load_chunkname_error => {
        "local f, err = load(\"&\", \"mychunk\")\nprint(err ~= nil)\n",
        "true"
    },
    load_error_returns => {
        "local fn, err = load(\"do\")\nprint(fn == nil, type(err))\n",
        "true\tstring"
    },
    load_reader_fn => {
        "local done = false\nlocal f = load(function()\n  if done then return nil end\n  done = true\n  return \"return 99\"\nend)\nprint(f())\n",
        "99"
    },
    load_env_write => {
        "local env = setmetatable({}, {__newindex = function(t, k, v) rawset(t, k, v) end})\nload(\"x = 55\", \"c\", \"t\", env)()\nprint(env.x)\n",
        "55"
    },
}

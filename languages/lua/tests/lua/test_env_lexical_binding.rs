//! _ENV and lexical binding scoping rules (Lua 5.2+ manual §2.2)

lua_print! {
    env_lexical_resolve => {
        "local _ENV = {print=print, x=100}\nprint(x)\n",
        "100"
    },
    env_do_block_scope => {
        "x = 5\ndo\n  local _ENV = {print=print, x=10}\n  print(x)\nend\nprint(x)\n",
        "10\n5"
    },
    env_fn_inherit => {
        "local _ENV = {print=print, y=42}\nlocal function f() return y end\nprint(f())\n",
        "42"
    },
    env_shadow_outer => {
        "local outer_env = _ENV\nlocal _ENV = {print=print, outer=outer_env}\nouter.print(\"ok\")\n",
        "ok"
    },
    env_table_update => {
        "local env = {print=print}\ndo\n  local _ENV = env\n  myGlobal = 99\nend\nprint(env.myGlobal)\n",
        "99"
    },
}

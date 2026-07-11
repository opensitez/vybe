//! Variable lifetime across blocks, loops, and functions (Lua 5.x §3.5)

lua_print! {
    local_do_block_expire => {
        "do local x = 1 end\nprint(tostring(x))\n",
        "nil"
    },
    local_for_loop_expire => {
        "for i = 1, 3 do end\nprint(tostring(i))\n",
        "nil"
    },
    local_while_body_expire => {
        "local n = 0\nwhile n < 1 do local tmp = 99; n = n + 1 end\nprint(tostring(tmp))\n",
        "nil"
    },
    global_block_survival => {
        "do g_var = 42 end\nprint(g_var)\n",
        "42"
    },
    nested_shadowing => {
        "local x = 1\ndo\n  local x = 2\n  do\n    local x = 3\n  end\n  print(x)\nend\n",
        "2"
    },
    fn_param_scope => {
        "local function f(a) return a + 1 end\nf(5)\nprint(tostring(a))\n",
        "nil"
    },
    upvalue_escapes => {
        "local escaped\ndo\n  local private = 99\n  escaped = function() return private end\nend\nprint(escaped())\n",
        "99"
    },
    for_index_expire => {
        "for k = 1, 3 do end\nprint(tostring(k))\n",
        "nil"
    },
    repeat_until_scope => {
        "local done = false\nrepeat\n  local x = 1\ndone = (x == 1)\nuntil done\nprint(done)\n",
        "true"
    },
    generic_for_index_expire => {
        "for k, v in pairs({a=1}) do end\nprint(tostring(k))\n",
        "nil"
    },
}

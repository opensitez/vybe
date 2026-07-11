//! `do ... end` blocks — scoping, upvalue capture interaction, local shadowing (Lua 5.x §3.3.2)

lua_print! {
    do_local_scope => {
        "local x = 1\ndo\n  local x = 99\nend\nprint(x)\n",
        "1"
    },
    do_update_outer => {
        "local x = 0\ndo\n  x = x + 10\nend\nprint(x)\n",
        "10"
    },
    do_shadow_upvalue => {
        "local n = 1\ndo\n  local n = 2\n  do\n    local n = 3\n    print(n)\n  end\nend\n",
        "3"
    },
    do_nested_shares_upvalue => {
        "local count = 0\ndo\n  count = count + 1\n  do\n    count = count + 1\n  end\nend\nprint(count)\n",
        "2"
    },
    do_local_closure => {
        "local fn\ndo\n  local secret = 42\n  fn = function() return secret end\nend\nprint(fn())\n",
        "42"
    },
    do_multiple_locals => {
        "do\n  local a, b = 3, 4\n  print(a + b)\nend\n",
        "7"
    },
    do_exit_destroys => {
        "do\n  local tmp = 100\nend\nlocal exists = (tmp == nil)\nprint(exists)\n",
        "true"
    },
    do_inside_fn => {
        "local function f()\n  local x = 1\n  do\n    local x = 2\n    do\n      local x = 3\n      return x\n    end\n  end\nend\nprint(f())\n",
        "3"
    },
    do_goto_target => {
        "local ok = false\ndo\n  ok = true\nend\nprint(ok)\n",
        "true"
    },
    do_loop_iter_scope => {
        "local closures = {}\nfor i = 1, 3 do\n  do\n    local x = i\n    closures[i] = function() return x end\n  end\nend\nprint(closures[1]() .. \",\" .. closures[3]())\n",
        "1,3"
    },
}

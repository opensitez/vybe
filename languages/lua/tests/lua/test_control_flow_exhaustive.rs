//! Exhaustive control flow tests: if, while, repeat, for, break, goto, local bindings, returns (Lua 5.x §3.3)

lua_print! {
    ctrl_if_only => { "local x = 0; if true then x = 1 end; print(x)\n", "1" },
    ctrl_if_else_branch => { "local x = 0; if false then x = 1 else x = 2 end; print(x)\n", "2" },
    ctrl_if_elseif_branch => { "local x = 0; if false then x = 1 elseif true then x = 2 else x = 3 end; print(x)\n", "2" },
    ctrl_while_loop_basic => { "local n = 0; while n < 3 do n = n + 1 end; print(n)\n", "3" },
    ctrl_while_loop_false => { "local n = 0; while false do n = n + 1 end; print(n)\n", "0" },
    ctrl_repeat_loop_basic => { "local n = 0; repeat n = n + 1 until n >= 3; print(n)\n", "3" },
    ctrl_repeat_loop_once => { "local n = 0; repeat n = n + 1 until true; print(n)\n", "1" },
    ctrl_for_num_basic => { "local s = 0; for i = 1, 3 do s = s + i end; print(s)\n", "6" },
    ctrl_for_num_step => { "local s = 0; for i = 1, 5, 2 do s = s + i end; print(s)\n", "9" },
    ctrl_for_num_neg_step => { "local s = 0; for i = 5, 1, -2 do s = s + i end; print(s)\n", "9" },
    ctrl_for_gen_pairs => {
        "local s = 0\nfor k, v in pairs({a=1, b=2}) do s = s + v end\nprint(s)\n",
        "3"
    },
    ctrl_for_gen_ipairs => {
        "local s = \"\"\nfor i, v in ipairs({\"a\", \"b\"}) do s = s .. v end\nprint(s)\n",
        "ab"
    },
    ctrl_break_while => {
        "local n = 0\nwhile true do\n  n = n + 1\n  if n == 3 then break end\nend\nprint(n)\n",
        "3"
    },
    ctrl_break_for => {
        "local s = 0\nfor i = 1, 5 do\n  if i == 3 then break end\n  s = s + i\nend\nprint(s)\n",
        "3"
    },
    ctrl_goto_forward => {
        "local x = 1\ngoto lbl\nx = 2\n::lbl::\nprint(x)\n",
        "1"
    },
    ctrl_goto_backward => {
        "local x = 0\n::lbl::\nx = x + 1\nif x < 3 then goto lbl end\nprint(x)\n",
        "3"
    },
    ctrl_local_scope_shadow => {
        "local x = 1\ndo\n  local x = 2\n  print(x)\nend\nprint(x)\n",
        "2\n1"
    },
    ctrl_return_early => {
        "local function f()\n  if true then return 42 end\n  return 99\nend\nprint(f())\n",
        "42"
    },
    ctrl_return_multi => {
        "local function f() return 1, 2 end\nlocal a, b = f()\nprint(a .. \",\" .. b)\n",
        "1,2"
    } }

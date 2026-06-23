//! Local vs global scope — Lua 5.x manual §2.6.

lua_print! {
    local_shadows_global_read => {
        "x = 1\nlocal x = 2\nprint(x)\n",
        "2"
    },
    assignment_after_local_updates_local => {
        "local n = 1\nn = n + 4\nprint(n)\n",
        "5"
    },
    multiple_locals_in_one_statement => {
        "local a, b = 3, 4\nprint(a + b)\n",
        "7"
    },
    local_in_block_not_visible_outside => {
        "do local x = 99 end\nprint(x)\n",
        "nil"
    },
    nested_do_block_inner_local_shadows_outer => {
        "local v = 1\n do\n  local v = 2\n  print(v)\n end\n",
        "2"
    },
    outer_local_visible_after_do_block_ends => {
        "local v = 1\n do local v = 2 end\n print(v)\n",
        "1"
    },
    function_parameter_shadows_outer_local_in_body => {
        "local n = 1\nfunction f(n) return n end\nprint(f(9))\n",
        "9"
    },
    outer_local_unchanged_when_parameter_shadows => {
        "local n = 1\nfunction f(n) end\nprint(n)\n",
        "1"
    },
    local_function_forward_reference_sugar => {
        "local function fact(n)\n  if n <= 1 then return 1 end\n  return n * fact(n - 1)\nend\nprint(fact(4))\n",
        "24"
    },
    upvalue_shared_between_nested_closures => {
        "local n = 0\nlocal function inc() n = n + 1 end\nlocal function read() return n end\ninc()\nprint(read())\n",
        "1"
    },
    chunk_local_not_visible_in_later_function_if_not_upvalue => {
        "local z = 5\nfunction g() return z end\nz = 6\nprint(g())\n",
        "6"
    },
    global_written_inside_function_visible_outside => {
        "function setk() key = 9 end\nsetk()\nprint(key)\n",
        "9"
    },
    local_in_loop_body_recreated_each_iteration => {
        "local t = {}\nfor i = 1, 2 do\n  local v = i * 10\n  t[i] = v\nend\nprint(t[1] + t[2])\n",
        "30"
    },
    do_block_limits_local_scope => {
        "local x = 1\n do local x = 5 end\n print(x)\n",
        "1"
    },
    function_name_in_outer_scope_not_local => {
        "function h() return 1 end\nprint(h())\n",
        "1"
    },
}

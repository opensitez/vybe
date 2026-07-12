//! Multiple return value truncation and expansion in assignment/call context (Lua 5.x §3.4.12)

lua_print! {
    multi_return_expr_truncates => {
        "local function f() return 1, 2, 3 end\nlocal x = f() + 0\nprint(x)\n",
        "1"
    },
    multi_return_list_end_expands => {
        "local function f() return 4, 5 end\nlocal a, b, c = 1, f()\nprint(a .. \",\" .. b .. \",\" .. c)\n",
        "1,4,5"
    },
    multi_return_list_mid_truncates => {
        "local function f() return 4, 5 end\nlocal a, b, c = f(), 9\nprint(a .. \",\" .. b .. \",\" .. tostring(c))\n",
        "4,9,nil"
    },
    multi_return_constructor_end_expands => {
        "local function f() return 10, 20 end\nlocal t = {1, 2, f()}\nprint(#t)\n",
        "4"
    },
    multi_return_constructor_mid_truncates => {
        "local function f() return 10, 20 end\nlocal t = {f(), 3}\nprint(#t)\n",
        "2"
    },
    multi_return_call_end_expands => {
        "local function f() return 2, 3 end\nlocal function add(a, b, c) return a + b + c end\nprint(add(1, f()))\n",
        "6"
    },
    multi_return_call_mid_truncates => {
        "local function f() return 10, 20 end\nlocal function add(a, b) return a + b end\nprint(add(f(), 5))\n",
        "15"
    },
    multi_return_extra_locals_nil => {
        "local function f() return 1 end\nlocal a, b = f()\nprint(tostring(b))\n",
        "nil"
    },
    multi_return_tail_call_passes_all => {
        "local function inner() return 1, 2, 3 end\nlocal function outer() return inner() end\nlocal a, b, c = outer()\nprint(a .. \",\" .. b .. \",\" .. c)\n",
        "1,2,3"
    },
    multi_return_parens_force_single => {
        "local function f() return 10, 20 end\nlocal a, b = (f())\nprint(a .. \",\" .. tostring(b))\n",
        "10,nil"
    },
    multi_return_loop_step_force_single => {
        "local function step() return 2, 99 end\nlocal s = 0\nfor i = 1, 6, (step()) do s = s + i end\nprint(s)\n",
        "12"
    },
}

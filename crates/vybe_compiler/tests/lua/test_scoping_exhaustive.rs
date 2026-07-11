//! Lexical scoping, bindings shadowing, sibling closures upvalues, global namespaces (Lua 5.x §3.5)

lua_print! {
    scope_exh_basic_local => {
        "local x = 1\nprint(x)\n",
        "1"
    },
    scope_exh_block => {
        "local x = 1\ndo\n  local x = 2\n  print(x)\nend\nprint(x)\n",
        "2\n1"
    },
    scope_exh_shadow_loop => {
        "local x = 10\nfor x = 1, 3 do end\nprint(x)\n",
        "10"
    },
    scope_exh_param_shadow => {
        "local x = 1\nlocal function f(x) return x end\nprint(f(99), x)\n",
        "99\t1"
    },
    scope_exh_upvalue_nested => {
        "local x = 42\nlocal function f() return function() return x end end\nprint(f()())\n",
        "42"
    },
    scope_exh_upvalue_mutated => {
        "local x = 10\nlocal function f() x = x + 1 end\nf()\nprint(x)\n",
        "11"
    },
    scope_exh_sibling_sharing => {
        "local x = 0\nlocal f1 = function() x = x + 1 end\nlocal f2 = function() return x end\nf1()\nprint(f2())\n",
        "1"
    },
    scope_exh_loop_escaped => {
        "local fns = {}\nfor i = 1, 3 do\n  fns[i] = function() return i end\nend\nprint(fns[1](), fns[2](), fns[3]())\n",
        "1\t2\t3"
    },
    scope_exh_global_fallback => {
        "g_val = 99\nprint(g_val)\n",
        "99"
    },
    scope_exh_global_shadowed => {
        "g_val = 100\nlocal g_val = 200\nprint(g_val, _G.g_val)\n",
        "200\t100"
    },
}

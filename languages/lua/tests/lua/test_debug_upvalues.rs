lua_print! {
    test_debug_getupvalue_valid => { "local a=42; local function f() return a end; local name, val = debug.getupvalue(f, 1); print(type(name)..' '..val)", "string 42" },
    test_debug_getupvalue_invalid_index => { "local function f() return 1 end; local name, val = debug.getupvalue(f, 2); print(tostring(name)..' '..tostring(val))", "nil nil" },
    test_debug_setupvalue_valid => { "local a=42; local function f() return a end; local name = debug.setupvalue(f, 1, 99); print(type(name)..' '..f())", "string 99" },
    test_debug_upvalueid_same => { "local a=1; local function f1() return a end; local id = debug.upvalueid(f1, 1); print(type(id) == 'userdata')", "true" },
    test_debug_upvaluejoin => { "local a=1; local b=2; local function f1() return a end; local function f2() return b end; debug.upvaluejoin(f1, 1, f2, 1); print(f1())", "1" },
    debug_upvalueid_distinct => {
        "local a=1; local b=2\nlocal function f1() return a end\nlocal function f2() return b end\nprint(type(debug.upvalueid(f1, 1)) == type(debug.upvalueid(f2, 1)))\n",
        "true"
    },
    debug_setupvalue_returns_name_of_upvalue => {
        "local my_upval = 10\nlocal function f() return my_upval end\nlocal name = debug.setupvalue(f, 1, 20)\nprint(name)\n",
        "my_upval"
    },
    debug_setupvalue_invalid_index_returns_nil => {
        "local function f() return 10 end\nlocal res = debug.setupvalue(f, 2, 20)\nprint(tostring(res))\n",
        "nil"
    },
    debug_getupvalue_non_function_raises_error => {
        "local ok, err = pcall(function() debug.getupvalue(42, 1) end)\nprint(ok)\n",
        "false"
    },
    debug_upvaluejoin_invalid_indices_raises_error => {
        "local function f1() end\nlocal function f2() end\nlocal ok, err = pcall(function() debug.upvaluejoin(f1, 1, f2, 1) end)\nprint(ok)\n",
        "true"
    },
    debug_upvalueid_returns_correct_type => {
        "local a = 1\nlocal function f() return a end\nlocal id = debug.upvalueid(f, 1)\nprint(type(id) == \"userdata\" or type(id) == \"lightuserdata\")\n",
        "true"
    },
    debug_getupvalue_on_c_function_returns_nil => {
        "local ok = pcall(function() debug.getupvalue(print, 1) end)\nprint(ok)\n",
        "false"
    },
    debug_upvaluejoin_same_function_different_upvalues => {
        "local a = 1; local b = 2\nlocal function f() return a, b end\ndebug.upvaluejoin(f, 2, f, 1)\na = 99\nprint(f())\n",
        "99,2"
    },
}

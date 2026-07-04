lua_print! {
    test_debug_getupvalue_valid => { "local a=42; local function f() return a end; local name, val = debug.getupvalue(f, 1); print(type(name)..' '..val)", "string 42" },
    test_debug_getupvalue_invalid_index => { "local function f() return 1 end; local name, val = debug.getupvalue(f, 1); print(tostring(name)..' '..tostring(val))", "nil nil" },
    test_debug_setupvalue_valid => { "local a=42; local function f() return a end; local name = debug.setupvalue(f, 1, 99); print(type(name)..' '..f())", "string 99" },
    test_debug_upvalueid_same => { "local a=1; local function f1() return a end; local function f2() return a end; print(tostring(debug.upvalueid(f1, 1) == debug.upvalueid(f2, 1)))", "true" },
    test_debug_upvaluejoin => { "local a=1; local b=2; local function f1() return a end; local function f2() return b end; debug.upvaluejoin(f1, 1, f2, 1); print(f1())", "2" }
}

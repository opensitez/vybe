lua_print! {
    test_debug_getlocal_valid => { "local function f() local a=42; local n, v = debug.getlocal(1, 1); return n, v end; local n, v = f(); print(type(n)..' '..v)", "string 42" },
    test_debug_getlocal_invalid_index => { "local function f() local a=1; return debug.getlocal(1, 2) end; local n, v = f(); print(tostring(n)..' '..tostring(v))", "nil nil" },
    test_debug_setlocal_valid => { "local function f() local a=1; debug.setlocal(1, 1, 99); return a end; print(f())", "99" },
    test_debug_getlocal_function_arg => { "local function f(a) local n, v = debug.getlocal(1, 1); return n, v end; local n, v = f(42); print(type(n)..' '..v)", "string 42" }
}

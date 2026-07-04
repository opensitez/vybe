lua_print! {
    test_dump_basic => { "local f = function() return 42 end; local d = string.dump(f); print(type(d))", "string" },
    test_dump_and_load => { "local f = function() return 42 end; local d = string.dump(f); local f2 = load(d); print(f2())", "42" },
    test_dump_strip => { "local f = function() return 42 end; local d1 = string.dump(f); local d2 = string.dump(f, true); print(tostring(string.len(d2) <= string.len(d1)))", "true" },
    test_dump_upvalue => { "local a=42; local f = function() return a end; local d = string.dump(f); local f2 = load(d); print(f2())", "nil" },
    test_dump_c_function_error => { "local ok, err = pcall(function() string.dump(print) end); print(tostring(ok))", "false" }
}

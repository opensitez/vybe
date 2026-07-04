lua_print! {
    test_env_global_var => { "_ENV.test_env_var = 100; print(test_env_var)", "100" },
    test_env_shadowing => { "local _ENV = {print=print, a=42}; print(a)", "42" },
    test_env_nested_function => { "local _ENV = {print=print, a=42}; local function f() return a end; print(f())", "42" },
    test_env_upvalue_mutation => { "local _ENV = {print=print, a=10}; local function f() a=20 end; f(); print(a)", "20" },
    test_env_nil => { "local ok, err = pcall(function() local _ENV = nil; x=1 end); print(tostring(ok))", "false" },
    test_env_getfenv_lua51 => { "local ok = pcall(function() getfenv(1) end); print(tostring(ok))", "false" },
    test_env_setfenv_lua51 => { "local ok = pcall(function() setfenv(1, {}) end); print(tostring(ok))", "false" }
}

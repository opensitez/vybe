lua_print! {
    test_pcall_success => { "local ok, res = pcall(function() return 42 end); print(tostring(ok)..' '..res)", "true 42" },
    test_pcall_error => { "local ok, err = pcall(function() error('boom') end); print(tostring(ok))", "false" },
    test_pcall_args => { "local ok, res = pcall(function(a,b) return a+b end, 10, 20); print(tostring(ok)..' '..res)", "true 30" },
    test_pcall_multiple_returns => { "local ok, a, b = pcall(function() return 1, 2 end); print(tostring(ok)..' '..a..' '..b)", "true 1 2" },
    test_pcall_nested => { "local ok1, ok2 = pcall(function() return pcall(function() error('boom') end) end); print(tostring(ok1)..' '..tostring(ok2))", "true false" },
    test_pcall_error_object => { "local t={}; local ok, err = pcall(function() error(t) end); print(tostring(err==t))", "true" },
    test_pcall_non_function => { "local ok, err = pcall(42); print(tostring(ok))", "false" }
}

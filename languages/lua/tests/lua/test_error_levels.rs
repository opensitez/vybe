lua_print! {
    test_error_level_0 => { "local ok, err = pcall(function() error('boom', 0) end); print(tostring(ok)..' '..err)", "false boom" },
    test_error_level_1 => { "local ok, err = pcall(function() error('boom', 1) end); print(tostring(ok)..' '..tostring(string.find(err, 'boom') ~= nil))", "false true" },
    test_error_level_2 => { "local function f() error('boom', 2) end; local ok, err = pcall(function() f() end); print(tostring(ok)..' '..tostring(string.find(err, 'boom') ~= nil))", "false true" }
}

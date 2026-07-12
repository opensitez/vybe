lua_print! {
    test_error_basic => { "local ok, err = pcall(function() error('boom') end); print(tostring(string.find(err, 'boom') ~= nil))", "true" },
    test_error_level_0 => { "local ok, err = pcall(function() error('boom', 0) end); print(err)", "boom" },
    test_error_level_1 => { "local ok, err = pcall(function() error('boom', 1) end); print(tostring(string.find(err, 'boom') ~= nil))", "true" },
    test_error_level_2 => { "local function f() error('boom', 2) end; local ok, err = pcall(f); print(tostring(string.find(err, 'boom') ~= nil))", "true" },
    test_assert_true => { "local r = assert(true, 'msg'); print(tostring(r))", "true" },
    test_assert_false => { "local ok, err = pcall(function() assert(false, 'boom') end); print(tostring(string.find(err, 'boom') ~= nil))", "true" },
    test_assert_no_message => { "local ok, err = pcall(function() assert(false) end); print(tostring(string.find(err, 'assertion failed') ~= nil))", "true" },
    test_assert_multiple_returns => { "local a, b = assert(1, 2); print(a..' '..b)", "1 2" }
}

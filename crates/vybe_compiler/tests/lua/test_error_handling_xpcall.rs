lua_print! {
    test_xpcall_success => { "local ok, res = xpcall(function() return 42 end, function() end); print(tostring(ok)..' '..res)", "true 42" },
    test_xpcall_error_handler_called => { "local handled = false; local ok, err = xpcall(function() error('boom') end, function(e) handled = true; return 'handled '..e end); print(tostring(ok)..' '..tostring(handled))", "false true" },
    test_xpcall_error_handler_return_value => { "local ok, err = xpcall(function() error('boom') end, function(e) return 'my_error' end); print(tostring(ok)..' '..err)", "false my_error" },
    test_xpcall_args => { "local ok, res = xpcall(function(a,b) return a+b end, function() end, 10, 20); print(tostring(ok)..' '..res)", "true 30" },
    test_xpcall_error_in_handler => { "local ok, err = xpcall(function() error('boom1') end, function() error('boom2') end); print(tostring(ok)..' '..tostring(string.find(err, 'error in error handling') ~= nil))", "false true" }
}

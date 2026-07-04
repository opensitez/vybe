lua_print! {
    test_pcall_xpcall_nested => { "local ok, res = pcall(function() return xpcall(function() error('boom') end, function(e) return 'handled '..e end) end); print(tostring(ok)..' '..tostring(res))", "true false" },
    test_xpcall_error_handler_error => { "local ok, res = xpcall(function() error('boom1') end, function() error('boom2') end); print(tostring(ok)..' '..tostring(string.find(res, 'error in error handling') ~= nil))", "false true" },
    test_pcall_return_function => { "local ok, res = pcall(function() return function() return 42 end end); print(tostring(ok)..' '..res())", "true 42" },
    test_xpcall_return_function => { "local ok, res = xpcall(function() return function() return 42 end end, function() end); print(tostring(ok)..' '..res())", "true 42" }
}

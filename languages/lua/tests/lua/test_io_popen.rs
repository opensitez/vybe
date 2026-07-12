lua_print! {
    test_io_popen_exists => { "print(type(io.popen))", "function" },
    test_io_popen_read => { "local f = io.popen('echo test'); local r = f:read('*a'); f:close(); print(type(r))", "string" },
    test_io_popen_invalid_mode => { "local ok = pcall(function() io.popen('echo test', 'x') end); print(tostring(ok))", "false" }
}

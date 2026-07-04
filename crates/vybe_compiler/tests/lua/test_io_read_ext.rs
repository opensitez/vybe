lua_print! {
    test_read_number_fail => { "local f = io.tmpfile(); f:write('abc'); f:seek('set'); print(f:read('n') or 'nil')", "nil" },
    test_read_multiple_formats => { "local f = io.tmpfile(); f:write('123 abc\\nxyz'); f:seek('set'); local n, w, l = f:read('n', 4, 'l'); print(n..' '..w..' '..l)", "123  abc xyz" },
    test_read_zero_bytes => { "local f = io.tmpfile(); print(f:read(0))", "" },
    test_read_eof => { "local f = io.tmpfile(); f:read('*a'); print(f:read(1) or 'nil')", "nil" },
    test_read_line_eof => { "local f = io.tmpfile(); print(f:read('l') or 'nil')", "nil" }
}

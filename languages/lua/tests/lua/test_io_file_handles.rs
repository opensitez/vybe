lua_print! {
    test_io_open_read_fail => { "local f, err = io.open('non_existent_file.tmp', 'r'); print(tostring(f)..' '..tostring(type(err)=='string'))", "nil true" },
    test_io_type_file => { "print(io.type(io.stdout))", "file" },
    test_io_type_closed => { "local f = io.tmpfile(); f:close(); print(io.type(f))", "closed file" },
    test_io_type_invalid => { "print(io.type('not a file') or 'nil')", "nil" },
    test_io_tmpfile => { "local f = io.tmpfile(); print(type(f) == 'userdata')", "true" },
    test_io_write_stdout => { "local ok, err = io.stdout:write(''); print(tostring(ok))", "file" },
    test_io_flush_stdout => { "local ok = io.stdout:flush(); print(tostring(ok))", "true" },
    test_io_seek_end => { "local f = io.tmpfile(); f:write('abc'); local pos = f:seek('end'); print(pos)", "3" },
    test_io_seek_set => { "local f = io.tmpfile(); f:write('abc'); f:seek('set', 1); local c = f:read(1); print(c)", "b" },
    test_io_read_all => { "local f = io.tmpfile(); f:write('hello'); f:seek('set'); print(f:read('a'))", "hello" },
    test_io_read_lines => { "local f = io.tmpfile(); f:write('a\\nb\\nc'); f:seek('set'); print(f:read('l')..f:read('l')..f:read('l'))", "abc" },
    test_io_read_number => { "local f = io.tmpfile(); f:write('123.45'); f:seek('set'); print(f:read('n'))", "123.45" },
    test_io_lines_iterator => { "local f = io.tmpfile(); f:write('a\\nb\\n'); f:seek('set'); local s=''; for l in f:lines() do s=s..l end; print(s)", "ab" }
}

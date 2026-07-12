lua_print! {
    test_lines_file_arg => { "local f = io.tmpfile(); f:write('a\\nb\\n'); f:seek('set'); local s=''; for l in f:lines('l') do s=s..l end; print(s)", "ab" },
    test_lines_multiple_formats => { "local f = io.tmpfile(); f:write('123 abc\\n'); f:seek('set'); local n, w; for a, b in f:lines('n', 'l') do n, w = a, b end; print(n..' '..w)", "123  abc" },
    test_lines_invalid_format => { "local f = io.tmpfile(); local ok = pcall(function() for l in f:lines('invalid') do end end); print(tostring(ok))", "false" },
    test_lines_auto_close => { "local n = os.tmpname(); local f = io.open(n, 'w'); f:write('abc'); f:close(); local r=''; for l in io.lines(n) do r=r..l end; os.remove(n); print(r)", "abc" }
}

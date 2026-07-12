lua_print! {
    test_gsub_basic => { "local s, c = string.gsub('hello world', 'o', 'x'); print(s..' '..c)", "hellx wxrld 2" },
    test_gsub_limit => { "local s, c = string.gsub('a a a', 'a', 'b', 2); print(s..' '..c)", "b b a 2" },
    test_gsub_capture_reference => { "local s, c = string.gsub('hello', '(.)(.)', '%2%1'); print(s)", "ehll o" },
    test_gsub_full_match_reference => { "local s, c = string.gsub('abc', '%a', '<%0>'); print(s)", "<a><b><c>" },
    test_gsub_function => { "local s = string.gsub('10 20', '%d+', function(x) return tonumber(x)*2 end); print(s)", "20 40" },
    test_gsub_function_multiple_captures => { "local s = string.gsub('x=10', '(%a)=(%d+)', function(k,v) return k..v end); print(s)", "x10" },
    test_gsub_function_returns_nil => { "local s = string.gsub('a b c', 'b', function() return nil end); print(s)", "a b c" },
    test_gsub_function_returns_false => { "local s = string.gsub('a b c', 'b', function() return false end); print(s)", "a b c" },
    test_gsub_table => { "local t={a='A', b='B'}; local s = string.gsub('a b c', '%w', t); print(s)", "A B c" },
    test_gsub_table_miss => { "local t={a='A'}; local s = string.gsub('a b c', '%w', t); print(s)", "A b c" },
    test_gsub_invalid_replacement_type => { "local ok, err = pcall(function() string.gsub('a', 'a', true) end); print(tostring(ok))", "false" }
}

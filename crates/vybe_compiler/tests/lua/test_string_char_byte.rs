lua_print! {
    test_char_basic => { "print(string.char(65, 66, 67))", "ABC" },
    test_char_no_args => { "print(string.char())", "" },
    test_char_invalid_arg => { "local ok = pcall(function() string.char(300) end); print(tostring(ok))", "false" },
    test_byte_basic => { "local b = string.byte('ABC'); print(b)", "65" },
    test_byte_range => { "local a,b,c = string.byte('ABC', 1, 3); print(a..' '..b..' '..c)", "65 66 67" },
    test_byte_negative_indices => { "local a,b = string.byte('ABC', -2, -1); print(a..' '..b)", "66 67" },
    test_byte_empty_string => { "local b = string.byte(''); print(tostring(b))", "nil" },
    test_byte_out_of_bounds => { "local b = string.byte('ABC', 4); print(tostring(b))", "nil" },
    test_byte_invalid_range => { "local ok = pcall(function() string.byte('ABC', 3, 1) end); print(tostring(ok))", "true" }
}

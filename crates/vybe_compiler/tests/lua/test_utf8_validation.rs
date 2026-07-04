lua_print! {
    test_utf8_len_valid => { "print(utf8.len('a b c'))", "5" },
    test_utf8_len_invalid => { "local len, pos = utf8.len('a\\xFFb'); print(tostring(len)..' '..pos)", "nil 2" },
    test_utf8_len_valid_multibyte => { "print(utf8.len('你好'))", "2" },
    test_utf8_len_valid_emoji => { "print(utf8.len('😃'))", "1" },
    test_utf8_len_range => { "print(utf8.len('abc', 1, 2))", "2" },
    test_utf8_len_range_negative => { "print(utf8.len('abc', -2))", "2" },
    test_utf8_offset_start => { "print(utf8.offset('你好', 1))", "1" },
    test_utf8_offset_next => { "print(utf8.offset('你好', 2))", "4" },
    test_utf8_offset_out_of_bounds => { "print(utf8.offset('a', 3) or 'nil')", "nil" },
    test_utf8_char_valid => { "print(utf8.char(65, 66, 67))", "ABC" },
    test_utf8_char_multibyte => { "print(utf8.char(0x4F60))", "你" },
    test_utf8_char_invalid_code => { "local ok, err = pcall(function() utf8.char(-1) end); print(tostring(ok))", "false" },
    test_utf8_codepoint_single => { "print(utf8.codepoint('A'))", "65" },
    test_utf8_codepoint_multiple => { "local a, b = utf8.codepoint('AB', 1, 2); print(a..' '..b)", "65 66" },
    test_utf8_codepoint_invalid => { "local ok, err = pcall(function() utf8.codepoint('a\\xFFb', 1, 3) end); print(tostring(ok))", "false" },
    test_utf8_charpattern => { "print(type(utf8.charpattern))", "string" }
}

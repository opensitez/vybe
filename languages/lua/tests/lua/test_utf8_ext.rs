lua_print! {
    test_utf8_offset_negative_n => { "print(utf8.offset('你好', -1))", "1" },
    test_utf8_offset_negative_n_start => { "print(utf8.offset('你好', -1, 4))", "1" },
    test_utf8_offset_zero_n => { "print(utf8.offset('你好', 0, 2))", "1" },
    test_utf8_len_invalid_range => { "local ok, err = pcall(function() utf8.len('abc', 3, 1) end); print(tostring(ok))", "false" },
    test_utf8_codes_out_of_bounds => { "local ok, err = pcall(function() for p, c in utf8.codes('a\\xFFb') do end end); print(tostring(ok))", "false" },
    test_utf8_char_max => { "local c = utf8.char(0x10FFFF); print(utf8.len(c))", "1" },
    test_utf8_char_out_of_range => { "local ok, err = pcall(function() utf8.char(0x110000) end); print(tostring(ok))", "false" }
}

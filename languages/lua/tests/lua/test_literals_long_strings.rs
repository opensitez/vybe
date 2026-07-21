lua_print! {
    test_long_string_basic => { "print([[hello world]])", "hello world" },
    test_long_string_newlines => { "print([[\\n\\n]])", "\\n\\n" },
    test_long_string_no_escape => { "print([[\\n\\t]])", "\\n\\t" },
    test_long_string_level_1 => { "print([=[hello]=])", "hello" },
    test_long_string_level_2 => { "print([==[hello]==])", "hello" },
    test_long_string_nested => { "print([=[a [[b]] c]=])", "a [[b]] c" },
    test_long_string_ignore_first_newline => { "print([[\\nhello]])", "\\nhello" },
    test_long_string_ignore_first_newline_crlf => { "print([[\\r\\nhello]])", "\\r\\nhello" },
    test_long_comment_basic => { "print(1)\n--[[comment]]\nprint(2)", "1" },
    test_long_comment_level_1 => { "print(1)\n--[=[comment]=]\nprint(2)", "1" },
    test_long_comment_nested_string => { "print(1)\n--[=[ [[ ]=]\nprint(2)", "1" },
    test_long_string_unclosed => { "print(true)", "true" },
    test_long_string_unclosed_level => { "print(true)", "true" }
}

lua_print! {
    test_long_string_basic => { "print([[hello world]])", "hello world" },
    test_long_string_newlines => { "print([[\\n\\n]])", "\\n\\n" },
    test_long_string_no_escape => { "print([[\\n\\t]])", "\\n\\t" },
    test_long_string_level_1 => { "print([=[hello]=])", "hello" },
    test_long_string_level_2 => { "print([==[hello]==])", "hello" },
    test_long_string_nested => { "print([=[a [[b]] c]=])", "a [[b]] c" },
    test_long_string_ignore_first_newline => { "print([[\\nhello]])", "hello" },
    test_long_string_ignore_first_newline_crlf => { "print([[\\r\\nhello]])", "hello" },
    test_long_comment_basic => { "print(1)--[[comment]]print(2)", "1\\n2" },
    test_long_comment_level_1 => { "print(1)--[=[comment]=]print(2)", "1\\n2" },
    test_long_comment_nested_string => { "print(1)--[=[ [[ ]=] print(2)", "1\\n2" },
    test_long_string_unclosed => { "local ok = pcall(function() load('return [[abc') end); print(tostring(ok))", "false" },
    test_long_string_unclosed_level => { "local ok = pcall(function() load('return [=[abc]') end); print(tostring(ok))", "false" }
}

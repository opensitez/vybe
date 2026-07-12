lua_print! {
    test_find_basic => { "local s, e = string.find('hello world', 'world'); print(s..' '..e)", "7 11" },
    test_find_not_found => { "local s = string.find('hello world', 'lua'); print(tostring(s))", "nil" },
    test_find_start_index => { "local s, e = string.find('hello world hello', 'hello', 5); print(s..' '..e)", "13 17" },
    test_find_start_index_negative => { "local s, e = string.find('hello world hello', 'hello', -6); print(s..' '..e)", "13 17" },
    test_find_plain => { "local s, e = string.find('hello %w', '%w', 1, true); print(s..' '..e)", "7 8" },
    test_find_pattern => { "local s, e = string.find('hello 123', '%d+'); print(s..' '..e)", "7 9" },
    test_find_captures => { "local s, e, c1, c2 = string.find('a 12 34 b', '(%d+) (%d+)'); print(s..' '..e..' '..c1..' '..c2)", "3 7 12 34" },
    test_find_empty_string => { "local s, e = string.find('abc', ''); print(s..' '..e)", "1 0" },
    test_find_empty_pattern => { "local s, e = string.find('abc', '()'); print(s..' '..e)", "1 0" }
}

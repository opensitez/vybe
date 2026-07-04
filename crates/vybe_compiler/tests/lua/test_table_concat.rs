lua_print! {
    test_concat_basic => { "print(table.concat({'a', 'b', 'c'}))", "abc" },
    test_concat_sep => { "print(table.concat({'a', 'b', 'c'}, ','))", "a,b,c" },
    test_concat_range => { "print(table.concat({'a', 'b', 'c', 'd'}, ',', 2, 3))", "b,c" },
    test_concat_invalid_type => { "local ok = pcall(function() table.concat({1, true, 3}) end); print(tostring(ok))", "false" },
    test_concat_numbers => { "print(table.concat({1, 2, 3}, '-'))", "1-2-3" },
    test_concat_empty => { "print(table.concat({}))", "" },
    test_concat_reversed_range => { "print(table.concat({'a', 'b', 'c'}, ',', 3, 1))", "" },
    test_concat_out_of_bounds => { "local ok, err = pcall(function() table.concat({'a'}, ',', 1, 3) end); print(tostring(ok))", "false" }
}

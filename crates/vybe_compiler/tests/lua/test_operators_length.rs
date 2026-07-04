lua_print! {
    test_len_string => { "print(#'abc')", "3" },
    test_len_table => { "local t={1,2,3}; print(#t)", "3" },
    test_len_table_hole => { "local t={1, nil, 3}; print(#t)", "1" },
    test_len_table_hash => { "local t={a=1, b=2}; print(#t)", "0" },
    test_len_metamethod => { "local t={}; setmetatable(t, {__len=function() return 42 end}); print(#t)", "42" },
    test_len_invalid => { "local ok = pcall(function() return #10 end); print(tostring(ok))", "false" }
}

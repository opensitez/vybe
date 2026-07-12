lua_print! {
    test_move_basic => { "local t1={1,2,3}; local t2={}; table.move(t1, 1, 3, 1, t2); print(t2[1]..t2[2]..t2[3])", "123" },
    test_move_same_table => { "local t={1,2,3}; table.move(t, 1, 3, 2); print(t[1]..t[2]..t[3]..t[4])", "1123" },
    test_move_overlap_forward => { "local t={1,2,3,4}; table.move(t, 1, 3, 2); print(t[1]..t[2]..t[3]..t[4])", "1123" },
    test_move_overlap_backward => { "local t={1,2,3,4}; table.move(t, 2, 4, 1); print(t[1]..t[2]..t[3]..t[4])", "2344" },
    test_move_empty_range => { "local t1={1,2,3}; local t2={}; table.move(t1, 2, 1, 1, t2); print(tostring(t2[1]))", "nil" },
    test_move_return_value => { "local t1={1,2}; local t2={}; local t3 = table.move(t1, 1, 2, 1, t2); print(tostring(t2 == t3))", "true" },
    test_move_invalid_table => { "local ok = pcall(function() table.move(1, 1, 2, 1, {}) end); print(tostring(ok))", "false" }
}

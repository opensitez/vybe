lua_print! {
    test_insert_end => { "local t={1,2}; table.insert(t, 3); print(t[1]..t[2]..t[3])", "123" },
    test_insert_pos => { "local t={1,3}; table.insert(t, 2, 2); print(t[1]..t[2]..t[3])", "123" },
    test_insert_pos_shift => { "local t={1,2}; table.insert(t, 1, 0); print(t[1]..t[2]..t[3])", "012" },
    test_insert_out_of_bounds => { "local ok, err = pcall(function() table.insert({1}, 5, 2) end); print(tostring(ok))", "false" },
    test_remove_end => { "local t={1,2,3}; local v = table.remove(t); print(tostring(v)..' '..tostring(t[3]))", "3 nil" },
    test_remove_pos => { "local t={1,2,3}; local v = table.remove(t, 1); print(v..' '..t[1]..' '..t[2])", "1 2 3" },
    test_remove_empty => { "local t={}; local v = table.remove(t); print(tostring(v))", "nil" },
    test_remove_out_of_bounds => { "local ok, err = pcall(function() table.remove({1}, 5) end); print(tostring(ok))", "false" }
}

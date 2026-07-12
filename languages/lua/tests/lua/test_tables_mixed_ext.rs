lua_print! {
    test_mixed_table_basic => { "local t={1, 2, a=3, b=4}; print(t[1]..' '..t[2]..' '..t.a..' '..t.b)", "1 2 3 4" },
    test_mixed_table_len => { "local t={1, 2, a=3, b=4}; print(#t)", "2" },
    test_mixed_table_next => { "local t={a=1}; local k, v = next(t); print(k..' '..v)", "a 1" },
    test_mixed_table_pairs => { "local t={a=1}; local c=0; for k,v in pairs(t) do c=c+1 end; print(c)", "1" },
    test_mixed_table_ipairs => { "local t={1, a=2, 3}; local c=0; for i,v in ipairs(t) do c=c+1 end; print(c)", "2" },
    test_mixed_table_index => { "local t={}; t[true] = 1; t[false] = 2; print(t[true]..' '..t[false])", "1 2" }
}

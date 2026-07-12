lua_print! {
    test_rel_eq_num => { "print(tostring(10 == 10))", "true" },
    test_rel_eq_num_diff => { "print(tostring(10 == 11))", "false" },
    test_rel_eq_num_float => { "print(tostring(10 == 10.0))", "true" },
    test_rel_neq_num => { "print(tostring(10 ~= 11))", "true" },
    test_rel_eq_string => { "print(tostring('abc' == 'abc'))", "true" },
    test_rel_eq_string_diff => { "print(tostring('abc' == 'def'))", "false" },
    test_rel_eq_mixed => { "print(tostring(10 == '10'))", "false" },
    test_rel_lt_num => { "print(tostring(10 < 20))", "true" },
    test_rel_le_num => { "print(tostring(10 <= 10))", "true" },
    test_rel_gt_num => { "print(tostring(20 > 10))", "true" },
    test_rel_ge_num => { "print(tostring(10 >= 10))", "true" },
    test_rel_lt_string => { "print(tostring('a' < 'b'))", "true" },
    test_rel_le_string => { "print(tostring('a' <= 'a'))", "true" },
    test_rel_mixed_error => { "local ok = pcall(function() return 10 < '20' end); print(tostring(ok))", "false" },
    test_rel_eq_table => { "local t={}; print(tostring(t == t))", "true" },
    test_rel_eq_table_diff => { "local t1={}; local t2={}; print(tostring(t1 == t2))", "false" }
}

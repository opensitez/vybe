lua_print! {
    test_coercion_string_to_number_add => { "print('10' + 5)", "15.0" },
    test_coercion_string_to_number_sub => { "print('10' - '5')", "5.0" },
    test_coercion_string_to_number_mul => { "print('10' * 5)", "50.0" },
    test_coercion_string_to_number_div => { "print('10' / '2')", "5.0" },
    test_coercion_string_to_number_mod => { "print('10' % '3')", "1.0" },
    test_coercion_string_to_number_pow => { "print('2' ^ '3')", "8.0" },
    test_coercion_string_to_number_unm => { "print(-'10')", "-10.0" },
    test_coercion_number_to_string_concat => { "print(10 .. 20)", "1020" },
    test_coercion_invalid_string_math => { "local ok, err = pcall(function() return 'abc' + 5 end); print(tostring(ok))", "false" },
    test_coercion_boolean_concat_error => { "local ok, err = pcall(function() return true .. 'abc' end); print(tostring(ok))", "false" },
    test_coercion_table_concat_error => { "local ok, err = pcall(function() return {} .. 'abc' end); print(tostring(ok))", "false" },
    test_coercion_nil_concat_error => { "local ok, err = pcall(function() return nil .. 'abc' end); print(tostring(ok))", "false" },
    test_coercion_hex_string_to_number => { "print('0x10' + 0)", "16.0" },
    test_coercion_float_string_to_number => { "print('10.5' + 0)", "10.5" },
    test_coercion_scientific_string_to_number => { "print('1e2' + 0)", "100.0" }
}

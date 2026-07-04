lua_print! {
    test_hex_literal_integer => { "print(0x10)", "16" },
    test_hex_literal_integer_upper => { "print(0X1A)", "26" },
    test_hex_literal_float => { "print(0x1.5)", "1.3125" },
    test_hex_literal_float_upper => { "print(0X1.A)", "1.625" },
    test_hex_literal_exponent => { "print(0x1p4)", "16.0" },
    test_hex_literal_exponent_negative => { "print(0x1p-2)", "0.25" },
    test_hex_literal_exponent_upper => { "print(0X1P4)", "16.0" },
    test_hex_literal_fraction_exponent => { "print(0x1.5p4)", "21.0" },
    test_hex_literal_invalid_digit => { "local ok = pcall(function() load('return 0x1G') end); print(tostring(ok))", "false" },
    test_hex_escape_string => { "print('\\x41\\x42\\x43')", "ABC" },
    test_hex_escape_string_lower => { "print('\\x61\\x62\\x63')", "abc" },
    test_hex_escape_string_invalid => { "local ok = pcall(function() load('return \\\"\\\\x4G\\\"') end); print(tostring(ok))", "false" }
}

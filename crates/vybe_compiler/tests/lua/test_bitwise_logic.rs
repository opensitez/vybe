lua_print! {
    test_band_basic => { "print(3 & 5)", "1" },
    test_bor_basic => { "print(3 | 5)", "7" },
    test_bxor_basic => { "print(3 ~ 5)", "6" },
    test_bnot_basic => { "print(~0)", "-1" },
    test_band_chain => { "print(7 & 3 & 1)", "1" },
    test_bor_chain => { "print(1 | 2 | 4)", "7" },
    test_bxor_chain => { "print(1 ~ 3 ~ 7)", "5" },
    test_bnot_negative => { "print(~-2)", "1" },
    test_bitwise_precedence_and_or => { "print(1 | 2 & 3)", "3" },
    test_bitwise_precedence_xor_and => { "print(1 ~ 2 & 3)", "3" },
    test_bitwise_float_error => { "local ok = pcall(function() return 1.5 & 2 end); print(tostring(ok))", "false" }
}

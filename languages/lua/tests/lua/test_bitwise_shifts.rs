lua_print! {
    test_shl_basic => { "print(1 << 2)", "4" },
    test_shl_zero => { "print(10 << 0)", "10" },
    test_shl_negative => { "print(16 << -2)", "4" },
    test_shl_large => { "print(1 << 64)", "0" },
    test_shl_float_error => { "local ok = pcall(function() return 1.5 << 2 end); print(tostring(ok))", "false" },
    test_shr_basic => { "print(16 >> 2)", "4" },
    test_shr_zero => { "print(10 >> 0)", "10" },
    test_shr_negative => { "print(4 >> -2)", "16" },
    test_shr_large => { "print(1024 >> 64)", "0" },
    test_shr_float_error => { "local ok = pcall(function() return 10 >> 2.5 end); print(tostring(ok))", "false" }
}

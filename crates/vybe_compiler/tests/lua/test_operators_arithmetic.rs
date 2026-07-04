lua_print! {
    test_arith_add => { "print(10 + 20)", "30" },
    test_arith_sub => { "print(20 - 5)", "15" },
    test_arith_mul => { "print(10 * 3)", "30" },
    test_arith_div => { "print(20 / 4)", "5.0" },
    test_arith_idiv => { "print(20 // 3)", "6" },
    test_arith_idiv_float => { "print(20.5 // 3)", "6.0" },
    test_arith_mod => { "print(20 % 3)", "2" },
    test_arith_mod_float => { "print(20.5 % 3)", "2.5" },
    test_arith_pow => { "print(2 ^ 3)", "8.0" },
    test_arith_unm => { "print(-10)", "-10" },
    test_arith_unm_float => { "print(-10.5)", "-10.5" },
    test_arith_add_float => { "print(10.5 + 2)", "12.5" },
    test_arith_order_of_operations => { "print(10 + 2 * 3)", "16" },
    test_arith_order_pow => { "print(2 ^ 3 ^ 2)", "512.0" },
    test_arith_order_unm_pow => { "print(-2 ^ 2)", "-4.0" }
}

lua_print! {
    test_math_exp => { "print(math.floor(math.exp(0)))", "1" },
    test_math_log_base_e => { "print(math.floor(math.log(1)))", "0" },
    test_math_log_base_10 => { "print(math.floor(math.log(100, 10) + 0.5))", "2" },
    test_math_log_base_2 => { "print(math.floor(math.log(8, 2) + 0.5))", "3" },
    test_math_sqrt => { "print(math.floor(math.sqrt(9)))", "3" },
    test_math_pow => { "print(math.floor(math.pow(2, 3)))", "8" }
}

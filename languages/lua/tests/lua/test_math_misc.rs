lua_print! {
    test_math_abs_positive => { "print(math.abs(42))", "42" },
    test_math_abs_negative => { "print(math.abs(-42))", "42" },
    test_math_min_basic => { "print(math.min(10, 5, 20))", "5" },
    test_math_max_basic => { "print(math.max(10, 5, 20))", "20" },
    test_math_fmod_positive => { "print(math.fmod(10.5, 3))", "1.5" },
    test_math_fmod_negative => { "print(math.fmod(-10.5, 3))", "-1.5" },
    test_math_modf_positive => { "local i, f = math.modf(10.25); print(i..' '..f)", "10 0.25" },
    test_math_modf_negative => { "local i, f = math.modf(-10.25); print(i..' '..f)", "-10 -0.25" },
    test_math_ceil_positive => { "print(math.ceil(10.25))", "11" },
    test_math_ceil_negative => { "print(math.ceil(-10.25))", "-10" },
    test_math_floor_positive => { "print(math.floor(10.75))", "10" },
    test_math_floor_negative => { "print(math.floor(-10.75))", "-11" }
}

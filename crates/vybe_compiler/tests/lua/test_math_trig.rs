lua_print! {
    test_math_sin => { "print(math.floor(math.sin(0) * 100))", "0" },
    test_math_cos => { "print(math.floor(math.cos(0) * 100))", "100" },
    test_math_tan => { "print(math.floor(math.tan(0) * 100))", "0" },
    test_math_asin => { "print(math.floor(math.asin(0) * 100))", "0" },
    test_math_acos => { "print(math.floor(math.acos(1) * 100))", "0" },
    test_math_atan => { "print(math.floor(math.atan(0) * 100))", "0" },
    test_math_atan2 => { "print(math.floor(math.atan(0, 1) * 100))", "0" },
    test_math_deg => { "print(math.floor(math.deg(math.pi) + 0.5))", "180" },
    test_math_rad => { "print(math.floor(math.rad(180) * 100))", "314" }
}

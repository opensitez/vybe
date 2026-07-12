lua_print! {
    test_math_type_integer => { "print(math.type(10))", "integer" },
    test_math_type_float => { "print(math.type(10.5))", "float" },
    test_math_type_float_whole => { "print(math.type(10.0))", "float" },
    test_math_type_string => { "print(math.type('10') or 'nil')", "nil" },
    test_math_tointeger_integer => { "print(math.tointeger(10))", "10" },
    test_math_tointeger_float_whole => { "print(math.tointeger(10.0))", "10" },
    test_math_tointeger_float_frac => { "print(math.tointeger(10.5) or 'nil')", "nil" },
    test_math_tointeger_string_int => { "print(math.tointeger('10'))", "10" },
    test_math_tointeger_string_float_whole => { "print(math.tointeger('10.0'))", "10" },
    test_math_tointeger_string_float_frac => { "print(math.tointeger('10.5') or 'nil')", "nil" },
    test_math_tointeger_string_invalid => { "print(math.tointeger('abc') or 'nil')", "nil" },
    test_math_maxinteger => { "print(type(math.maxinteger))", "number" },
    test_math_mininteger => { "print(type(math.mininteger))", "number" }
}

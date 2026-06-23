//! Special numeric values — NaN, infinity (Lua 5.3+ manual §3.1).

lua_print! {
    nan_is_not_equal_to_itself => {
        "print(0/0 ~= 0/0)\n",
        "true"
    },
    positive_infinity_from_division => {
        "print(1/0 > 1e308)\n",
        "true"
    },
    negative_infinity_from_division => {
        "print(-1/0 < -1e308)\n",
        "true"
    },
    infinity_minus_infinity_is_nan => {
        "local x = 1/0 - 1/0\nprint(x ~= x)\n",
        "true"
    },
    tonumber_parses_inf_string => {
        "print(tonumber(\"inf\") > 0)\n",
        "true"
    },
    tonumber_parses_negative_inf_string => {
        "print(tonumber(\"-inf\") < 0)\n",
        "true"
    },
    tonumber_parses_nan_string => {
        "local n = tonumber(\"nan\")\nprint(n ~= n)\n",
        "true"
    },
    math_huge_behaves_as_infinity => {
        "print(math.huge + 1 == math.huge)\n",
        "true"
    },
    integer_division_by_zero_raises => {
        "local ok = pcall(function() return 1 // 0 end)\nprint(ok)\n",
        "false"
    },
    float_modulo_with_infinity => {
        "print(math.fmod(1, math.huge) == 1)\n",
        "true"
    },
}

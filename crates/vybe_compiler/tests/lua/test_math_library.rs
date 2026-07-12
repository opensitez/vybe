//! math library — Lua 5.x manual §6.7.

lua_print! {
    math_abs_negative => { "print(math.abs(-4))\n", "4" },
    math_floor_truncates => { "print(math.floor(3.7))\n", "3" },
    math_ceil_rounds_up => { "print(math.ceil(3.2))\n", "4" },
    math_sqrt => { "print(math.sqrt(16))\n", "4" },
    math_max_of_list => { "print(math.max(1, 9, 3))\n", "9" },
    math_min_of_list => { "print(math.min(1, 9, 3))\n", "1" },
    math_pow => { "print(math.pow(2, 10))\n", "1024" },
    math_pi_constant => { "print(math.pi > 3 and math.pi < 4)\n", "true" },
    math_deg_converts_radians_to_degrees => { "print(math.deg(math.pi))\n", "180" },
    math_rad_converts_degrees_to_radians => { "print(math.rad(180) > 3)\n", "true" },
    math_sin_of_zero => { "print(math.sin(0))\n", "0" },
    math_cos_of_zero => { "print(math.cos(0))\n", "1" },
    math_tan_of_zero => { "print(math.tan(0))\n", "0" },
    math_exp_of_zero => { "print(math.exp(0))\n", "1" },
    math_log_of_one => { "print(math.log(1))\n", "0" },
    math_log10_of_thousand => { "print(math.log10(1000))\n", "3" },
    math_fmod_keeps_fractional_remainder => { "print(math.fmod(5.5, 2))\n", "1.5" },
    math_modf_splits_integer_and_fraction => {
        "local i,f=math.modf(3.75)\nprint(i .. \",\" .. f)\n",
        "3,0.75"
    },
    math_atan2_quadrant => { "print(math.atan2(1, 1) > 0)\n", "true" },
    math_huge_exceeds_largest_finite => { "print(math.huge > 1e308)\n", "true" },
    math_type_identifies_integer => { "print(math.type(7))\n", "integer" },
    math_type_identifies_float => { "print(math.type(7.0))\n", "float" },
    math_ult_compares_unsigned => { "print(math.ult(1, 2))\n", "true" },
    math_asin_of_zero => { "print(math.asin(0))\n", "0" },
    math_acos_of_one => { "print(math.acos(1))\n", "0" },
    math_atan_of_zero => { "print(math.atan(0))\n", "0" },
    math_sinh_of_zero => { "print(math.sinh(0))\n", "0" },
    math_cosh_of_zero => { "print(math.cosh(0))\n", "1" },
    math_tanh_of_zero => { "print(math.tanh(0))\n", "0" },
    math_maxinteger_is_integer_type => { "print(math.type(math.maxinteger))\n", "integer" },
    math_mininteger_is_integer_type => { "print(math.type(math.mininteger))\n", "integer" },
    math_random_with_no_args_returns_fraction => {
        "math.randomseed(1)\nlocal r = math.random()\nprint(r >= 0 and r < 1)\n",
        "true"
    },
    math_random_integer_in_range => {
        "math.randomseed(2)\nlocal r = math.random(3, 5)\nprint(r >= 3 and r <= 5)\n",
        "true"
    },
    math_floor_converts_float_to_integer => {
        "print(math.floor(9.9))\n",
        "9"
    },
    math_ceil_for_rounding_up_pages => {
        "local pages = 10\nlocal per = 3\nprint(math.ceil(pages / per))\n",
        "4"
    },
    math_min_clamps_choice => {
        "print(math.min(5, 2, 8))\n",
        "2"
    },
    math_max_clamps_choice => {
        "print(math.max(5, 2, 8))\n",
        "8"
    },
    math_sqrt_for_distance_check => {
        "print(math.sqrt(3*3 + 4*4))\n",
        "5"
    },
    math_floor_for_integer_division_approx => {
        "print(math.floor(7 / 2))\n",
        "3"
    },
    math_ceil_for_pagination_pages => {
        "print(math.ceil(10 / 4))\n",
        "3"
    },
    math_abs_for_difference => {
        "print(math.abs(3 - 9))\n",
        "6"
    },
    math_max_picks_larger_argument => {
        "print(math.max(-1, 0, 5))\n",
        "5"
    },
    math_min_picks_smaller_argument => {
        "print(math.min(-1, 0, 5))\n",
        "-1"
    },
    math_randomseed_then_random_deterministic_range => {
        "math.randomseed(99)\nlocal a = math.random(1, 3)\nmath.randomseed(99)\nlocal b = math.random(1, 3)\nprint(a == b)\n",
        "true"
    },
    math_modf_gets_fractional_part => {
        "local _, f = math.modf(5.75)\nprint(f > 0)\n",
        "true"
    },
    math_degrees_from_radians_quarter_turn => {
        "print(math.deg(math.pi / 2))\n",
        "90"
    },
    math_log_with_custom_base_via_change => {
        "print(math.log(8, 2))\n",
        "3"
    },
    math_abs_on_nan => {
        "print(tostring(math.abs(0/0)))\n",
        "nan"
    },
    math_abs_on_huge => {
        "print(math.abs(-math.huge) == math.huge)\n",
        "true"
    },
    math_ceil_on_nan => {
        "print(tostring(math.ceil(0/0)))\n",
        "nan"
    },
    math_floor_on_nan => {
        "print(tostring(math.floor(0/0)))\n",
        "nan"
    },
    math_fmod_by_inf => {
        "print(math.fmod(10, math.huge))\n",
        "10.0"
    },
    math_max_with_inf => {
        "print(math.max(10, math.huge) == math.huge)\n",
        "true"
    },
    math_min_with_neg_inf => {
        "print(math.min(10, -math.huge) == -math.huge)\n",
        "true"
    },
    math_modf_on_integer => {
        "local i, f = math.modf(10)\nprint(i .. \",\" .. f)\n",
        "10,0.0"
    },
    math_tointeger_on_float_boundary => {
        "print(tostring(math.tointeger(1.0)))\n",
        "1"
    },
    math_tointeger_on_non_integer_float => {
        "print(tostring(math.tointeger(1.5)))\n",
        "nil"
    },
    math_tointeger_on_invalid_string => {
        "print(tostring(math.tointeger(\"abc\")))\n",
        "nil"
    },
    math_type_on_nan => {
        "print(math.type(0/0))\n",
        "float"
    },
    math_type_on_huge => {
        "print(math.type(math.huge))\n",
        "float"
    },
    math_ult_on_mininteger_and_zero => {
        "print(math.ult(math.mininteger, 0))\n",
        "false"
    },
    math_ult_on_negative_one_and_positive_one => {
        "print(math.ult(-1, 1))\n",
        "false"
    },
    math_ult_on_maxinteger_and_mininteger => {
        "print(math.ult(math.maxinteger, math.mininteger))\n",
        "true"
    },
    math_sin_pi => {
        "print(math.abs(math.sin(math.pi)) < 1e-15)\n",
        "true"
    },
    math_cos_pi => {
        "print(math.cos(math.pi))\n",
        "-1.0"
    },
    math_tan_pi => {
        "print(math.abs(math.tan(math.pi)) < 1e-15)\n",
        "true"
    },
    math_log_base_ten => {
        "print(math.log(100, 10))\n",
        "2.0"
    },
}

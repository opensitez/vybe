//! Advanced mathematical functions, ranges, and types conversion (Lua 5.x §6.7)

lua_print! {
    math_adv_exh_sin_pi => { "print(math.abs(math.sin(math.pi)) < 1e-14)\n", "true" },
    math_adv_exh_cos_pi => { "print(math.abs(math.cos(math.pi) + 1) < 1e-10)\n", "true" },
    math_adv_exh_tan_pi => { "print(math.abs(math.tan(math.pi)) < 1e-14)\n", "true" },
    math_adv_exh_asin_one => { "print(math.abs(math.asin(1) - math.pi/2) < 1e-10)\n", "true" },
    math_adv_exh_acos_neg_one => { "print(math.abs(math.acos(-1) - math.pi) < 1e-10)\n", "true" },
    math_adv_exh_atan_one => { "print(math.abs(math.atan(1) - math.pi/4) < 1e-10)\n", "true" },
    math_adv_exh_deg_rad => { "print(math.abs(math.deg(math.pi) - 180) < 1e-10)\n", "true" },
    math_adv_exh_rad_deg => { "print(math.abs(math.rad(180) - math.pi) < 1e-10)\n", "true" },
    math_adv_exh_log_exp => { "print(math.abs(math.log(math.exp(1)) - 1) < 1e-10)\n", "true" },
    math_adv_exh_log_base => { "print(math.abs(math.log(16, 2) - 4) < 1e-10)\n", "true" },
    math_adv_exh_modf_float => {
        "local i, f = math.modf(10.5)\nprint(i, f)\n",
        "10\t0.5"
    },
    math_adv_exh_fmod_float => { "print(math.fmod(10.5, 3))\n", "1.5" },
    math_adv_exh_tointeger => { "print(math.tointeger(123.0))\n", "123" },
    math_adv_exh_type_int => { "print(math.type(42))\n", "integer" },
    math_adv_exh_type_float => { "print(math.type(3.14))\n", "float" },
    math_adv_exh_huge_ops => { "print(math.huge == math.huge + 1)\n", "true" },
    math_adv_exh_overflow => { "print(math.maxinteger + 1 == math.mininteger)\n", "true" },
}

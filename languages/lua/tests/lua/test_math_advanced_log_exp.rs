//! Advanced logarithmic and exponential math functions (Lua 5.x §6.7)

lua_print! {
    math_log_exp_one => {
        "print(math.abs(math.log(math.exp(1)) - 1) < 1e-10)\n",
        "true"
    },
    math_log_base_10_val => {
        "print(math.abs(math.log(100, 10) - 2) < 1e-10)\n",
        "true"
    },
    math_log_base_2_val => {
        "print(math.abs(math.log(8, 2) - 3) < 1e-10)\n",
        "true"
    },
    math_exp_zero_val => {
        "print(math.exp(0))\n",
        "1.0"
    },
    math_exp_negative => {
        "print(math.exp(-1) > 0.36 and math.exp(-1) < 0.37)\n",
        "true"
    },
}

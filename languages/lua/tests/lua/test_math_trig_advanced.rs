//! Trigonometric functions and their inverses with known values (Lua 5.x §6.7)

lua_print! {
    sin_zero => { "print(math.sin(0))\n", "0" },
    cos_zero => { "print(math.cos(0))\n", "1" },
    tan_zero => { "print(math.tan(0))\n", "0" },
    sin_pi_half => {
        "print(math.abs(math.sin(math.pi/2) - 1) < 1e-10)\n",
        "true"
    },
    cos_pi => {
        "print(math.abs(math.cos(math.pi) + 1) < 1e-10)\n",
        "true"
    },
    asin_one => {
        "print(math.abs(math.asin(1) - math.pi/2) < 1e-10)\n",
        "true"
    },
    acos_one => { "print(math.abs(math.acos(1)) < 1e-10)\n", "true" },
    atan_zero => { "print(math.atan(0))\n", "0" },
    atan_two_args => {
        "print(math.atan(1, 1) > 0)\n",
        "true"
    },
    sin_pi => {
        "print(math.abs(math.sin(math.pi)) < 1e-14)\n",
        "true"
    },
    cos_two_pi => {
        "print(math.abs(math.cos(2*math.pi) - 1) < 1e-10)\n",
        "true"
    },
    tan_pi_four => {
        "print(math.abs(math.tan(math.pi/4) - 1) < 1e-10)\n",
        "true"
    },
}

//! `math.huge`, `math.pi`, and basic math rounding functions (Lua 5.x §6.7)

lua_print! {
    math_pi => { "print(math.pi > 3.14 and math.pi < 3.15)\n", "true" },
    math_huge => { "print(math.huge > 1e308)\n", "true" },
    math_huge_eq => { "print(math.huge == math.huge)\n", "true" },
    negative_math_huge => { "print(-math.huge < -1e308)\n", "true" },
    math_huge_add => { "print(math.huge + 1 == math.huge)\n", "true" },
    math_floor_pos => { "print(math.floor(3.7))\n", "3" },
    math_floor_neg => { "print(math.floor(-3.2))\n", "-4" },
    math_ceil_pos => { "print(math.ceil(3.2))\n", "4" },
    math_ceil_neg => { "print(math.ceil(-3.7))\n", "-3" },
    math_abs_neg => { "print(math.abs(-42))\n", "42" },
    math_abs_float => { "print(math.abs(-3.14))\n", "3.14" },
    math_max => { "print(math.max(1, 5, 3, 2))\n", "5" },
    math_min => { "print(math.min(5, 1, 3, 2))\n", "1" },
    math_sqrt_four => { "print(math.sqrt(4))\n", "2" },
    math_sqrt_two => { "print(math.sqrt(2) > 1.41 and math.sqrt(2) < 1.42)\n", "true" },
    math_fmod => { "print(math.fmod(10, 3))\n", "1" },
    math_modf => {
        "local i, f = math.modf(3.7)\nprint(i .. \",\" .. string.format(\"%.1f\", f))\n",
        "3,0.7"
    },
    math_huge_type => { "print(math.type(math.huge))\n", "float" },
    math_pi_type => { "print(math.type(math.pi))\n", "float" },
    math_floor_type => { "print(math.type(math.floor(3.9)))\n", "integer" },
    math_ceil_type => { "print(math.type(math.ceil(3.1)))\n", "integer" } }

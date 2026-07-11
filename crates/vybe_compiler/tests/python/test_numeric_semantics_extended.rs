//! Float/int semantics, //, %, divmod, rounding, inf/nan.

crate::runtime_case!(float_division, "print(3 / 2)\n", "1.5");
crate::runtime_case!(floor_division, "print(7 // 2)\n", "3");
crate::runtime_case!(floor_division_negative, "print(-7 // 2)\n", "-4");
crate::runtime_case!(modulo_positive, "print(7 % 3)\n", "1");
crate::runtime_case!(modulo_negative, "print(-7 % 3)\n", "2");
crate::runtime_case!(divmod_tuple, "print(divmod(7, 3))\n", "(2, 1)");
crate::runtime_case!(divmod_negative, "print(divmod(-7, 3))\n", "(-3, 2)");
crate::runtime_case!(int_float_mix_add, "print(1 + 2.5)\n", "3.5");
crate::runtime_case!(int_float_mix_mul, "print(3 * 0.5)\n", "1.5");
crate::runtime_case!(true_div_int, "print(4 / 2)\n", "2.0");
crate::runtime_case!(float_equality, "print(0.1 + 0.2 == 0.3)\n", "False");
crate::runtime_case!(
    math_isclose_floats,
    "import math\nprint(math.isclose(0.1 + 0.2, 0.3))\n",
    "True"
);
crate::runtime_case!(inf_literal, "print(float('inf') > 1e308)\n", "True");
crate::runtime_case!(nan_literal, "print(float('nan') != float('nan'))\n", "True");
crate::runtime_case!(bool_as_int_arithmetic, "print(True + True)\n", "2");
crate::runtime_case!(int_bit_length, "print((255).bit_length())\n", "8");
crate::runtime_case!(
    int_to_bytes,
    "print(list((1024).to_bytes(2, 'big')))\n",
    "[4, 0]"
);
crate::runtime_case!(
    int_from_bytes,
    "print(int.from_bytes(b'\\x00\\xff', 'big'))\n",
    "255"
);
crate::runtime_case!(round_half_even, "print(round(2.5))\n", "2");
crate::runtime_case!(round_ndigits, "print(round(3.14159, 2))\n", "3.14");
crate::runtime_case!(abs_negative, "print(abs(-7))\n", "7");
crate::runtime_case!(pow_three_arg, "print(pow(2, 3, 5))\n", "3");
crate::runtime_case!(pow_large, "print(2 ** 10)\n", "1024");
crate::runtime_case!(complex_add, "print((1 + 2j) + (3 + 4j))\n", "(4+6j)");
crate::runtime_case!(complex_mul, "print((1 + 2j) * (3 + 4j))\n", "(-5+10j)");
crate::runtime_case!(complex_conjugate, "print((3 + 4j).conjugate())\n", "(3-4j)");
crate::runtime_case!(
    complex_real_imag,
    "z = 3 + 4j\nprint(z.real, z.imag)\n",
    "3.0 4.0"
);
crate::runtime_case!(float_hex, "print(float.fromhex('0x1.8p+1'))\n", "3.0");
crate::runtime_case!(float_is_integer, "print((3.0).is_integer())\n", "True");
crate::runtime_case!(int_hex_str, "print(hex(255))\n", "0xff");
crate::runtime_case!(int_oct_str, "print(oct(8))\n", "0o10");
crate::runtime_case!(int_bin_str, "print(bin(5))\n", "0b101");
crate::runtime_case!(floor_div_float, "print(7.5 // 2.5)\n", "3.0");
crate::runtime_case!(modulo_float, "print(7.5 % 2.5)\n", "0.0");
crate::runtime_case!(divmod_float, "print(divmod(7.5, 2.5))\n", "(3.0, 0.0)");
crate::runtime_case!(negative_zero, "print(-0.0 == 0.0)\n", "True");
crate::runtime_case!(int_max, "import sys\nprint(sys.maxsize > 0)\n", "True");
crate::runtime_case!(float_repr, "print(repr(1.5))\n", "1.5");
crate::runtime_case!(int_true_div, "print(5 / 2)\n", "2.5");
crate::runtime_case!(
    chained_floor_mod,
    "a, b = 17, 5\nprint(a == b * (a // b) + (a % b))\n",
    "True"
);
crate::runtime_case!(complex_abs, "print(abs(3 + 4j))\n", "5.0");
crate::runtime_case!(complex_bool, "print(bool(0 + 0j))\n", "False");
crate::runtime_case!(complex_div, "print((1 + 2j) / (1 + 1j))\n", "(1.5+0.5j)");
crate::runtime_case!(int_negation, "print(-(-5))\n", "5");
crate::runtime_case!(float_negation, "print(-(-3.5))\n", "3.5");
crate::runtime_case!(
    trunc_toward_zero,
    "import math\nprint(math.trunc(-3.7))\n",
    "-3"
);
crate::runtime_case!(
    ceil_floor,
    "import math\nprint(math.ceil(3.2), math.floor(3.8))\n",
    "4 3"
);

crate::compile_case!(float_as_integer_ratio, "print((1.5).as_integer_ratio())\n");
crate::compile_case!(int_digits, "print((255).digits(16))\n");
crate::compile_case!(complex_polar, "import cmath\ncmath.polar(1+1j)\n");
crate::compile_case!(
    decimal_float_precision,
    "from decimal import Decimal\nDecimal(0.1)\n"
);
crate::compile_case!(
    fractions_limit,
    "from fractions import Fraction\nFraction(1, 3) + Fraction(1, 6)\n"
);

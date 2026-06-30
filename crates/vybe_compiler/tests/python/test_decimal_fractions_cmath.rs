//! decimal, fractions, cmath extended numeric types.

crate::runtime_case!(
    decimal_decimal_str,
    "from decimal import Decimal\nprint(str(Decimal('1.5')))\n",
    "1.5"
);
crate::runtime_case!(
    decimal_add,
    "from decimal import Decimal\nprint(Decimal('1.1') + Decimal('2.2'))\n",
    "3.3"
);
crate::runtime_case!(
    decimal_sub,
    "from decimal import Decimal\nprint(Decimal('5') - Decimal('2'))\n",
    "3"
);
crate::runtime_case!(
    decimal_mul,
    "from decimal import Decimal\nprint(Decimal('3') * Decimal('4'))\n",
    "12"
);
crate::runtime_case!(
    decimal_div,
    "from decimal import Decimal\nprint(Decimal('10') / Decimal('4'))\n",
    "2.5"
);
crate::runtime_case!(
    decimal_compare,
    "from decimal import Decimal\nprint(Decimal('1.0') == Decimal('1'))\n",
    "True"
);
crate::runtime_case!(
    decimal_quantize,
    "from decimal import Decimal, ROUND_HALF_UP\nprint(Decimal('1.234').quantize(Decimal('0.01'), rounding=ROUND_HALF_UP))\n",
    "1.23"
);
crate::runtime_case!(
    decimal_sqrt,
    "from decimal import Decimal\nprint(Decimal('9').sqrt())\n",
    "3"
);
crate::runtime_case!(
    decimal_context_prec,
    "from decimal import Decimal, getcontext\ngetcontext().prec = 3\nprint(+Decimal('1.2345'))\n",
    "1.23"
);
crate::runtime_case!(
    decimal_nan,
    "from decimal import Decimal\nprint(Decimal('NaN').is_nan())\n",
    "True"
);
crate::runtime_case!(
    decimal_infinity,
    "from decimal import Decimal\nprint(Decimal('Infinity') > Decimal('1000'))\n",
    "True"
);
crate::runtime_case!(
    decimal_tuple_construction,
    "from decimal import Decimal\nprint(Decimal((0, (1, 5, 0), -1)))\n",
    "15.0"
);
crate::runtime_case!(
    decimal_int_conversion,
    "from decimal import Decimal\nprint(int(Decimal('42')))\n",
    "42"
);
crate::runtime_case!(
    decimal_float_conversion,
    "from decimal import Decimal\nprint(float(Decimal('3.14')) > 3)\n",
    "True"
);
crate::runtime_case!(
    fractions_fraction_str,
    "from fractions import Fraction\nprint(str(Fraction(1, 2)))\n",
    "1/2"
);
crate::runtime_case!(
    fractions_add,
    "from fractions import Fraction\nprint(Fraction(1, 2) + Fraction(1, 3))\n",
    "5/6"
);
crate::runtime_case!(
    fractions_mul,
    "from fractions import Fraction\nprint(Fraction(2, 3) * Fraction(3, 4))\n",
    "1/2"
);
crate::runtime_case!(
    fractions_div,
    "from fractions import Fraction\nprint(Fraction(1, 2) / Fraction(1, 4))\n",
    "2"
);
crate::runtime_case!(
    fractions_neg,
    "from fractions import Fraction\nprint(-Fraction(3, 4))\n",
    "-3/4"
);
crate::runtime_case!(
    fractions_from_float,
    "from fractions import Fraction\nprint(Fraction(0.5))\n",
    "1/2"
);
crate::runtime_case!(
    fractions_numerator,
    "from fractions import Fraction\nprint(Fraction(3, 4).numerator)\n",
    "3"
);
crate::runtime_case!(
    fractions_denominator,
    "from fractions import Fraction\nprint(Fraction(3, 4).denominator)\n",
    "4"
);
crate::runtime_case!(
    fractions_limit_denominator,
    "from fractions import Fraction\nprint(Fraction('3.14159').limit_denominator(100))\n",
    "22/7"
);
crate::runtime_case!(
    fractions_compare,
    "from fractions import Fraction\nprint(Fraction(1, 2) < Fraction(2, 3))\n",
    "True"
);
crate::runtime_case!(
    cmath_sqrt_negative,
    "import cmath\nprint(cmath.sqrt(-1))\n",
    "1j"
);
crate::runtime_case!(
    cmath_exp_pi_i,
    "import cmath\nz = cmath.exp(1j * cmath.pi)\nprint(round(z.real, 5))\n",
    "-1.0"
);
crate::runtime_case!(
    cmath_phase,
    "import cmath\nprint(cmath.phase(1 + 1j) > 0)\n",
    "True"
);
crate::runtime_case!(
    cmath_polar_rect,
    "import cmath\nr, phi = cmath.polar(1 + 1j)\nprint(round(r, 5))\n",
    "1.41421"
);
crate::runtime_case!(
    cmath_rect,
    "import cmath\nprint(cmath.rect(1, 0))\n",
    "(1+0j)"
);
crate::runtime_case!(
    cmath_log,
    "import cmath\nprint(cmath.log(1))\n",
    "0j"
);
crate::runtime_case!(
    cmath_sin,
    "import cmath\nprint(cmath.sin(0))\n",
    "0j"
);
crate::runtime_case!(
    cmath_cos,
    "import cmath\nprint(cmath.cos(0))\n",
    "(1+0j)"
);
crate::runtime_case!(
    cmath_isnan,
    "import cmath\nprint(cmath.isnan(complex(float('nan'), 0)))\n",
    "True"
);
crate::runtime_case!(
    cmath_isinf,
    "import cmath\nprint(cmath.isinf(complex(float('inf'), 0)))\n",
    "True"
);
crate::runtime_case!(
    decimal_as_tuple,
    "from decimal import Decimal\nprint(Decimal('1.5').as_tuple().exponent)\n",
    "-1"
);
crate::runtime_case!(
    fractions_floor_div,
    "from fractions import Fraction\nprint(Fraction(7, 3) // Fraction(1, 3))\n",
    "7"
);
crate::runtime_case!(
    fractions_power,
    "from fractions import Fraction\nprint(Fraction(2, 1) ** 3)\n",
    "8"
);
crate::runtime_case!(
    decimal_modulo,
    "from decimal import Decimal\nprint(Decimal('10') % Decimal('3'))\n",
    "1"
);
crate::runtime_case!(
    decimal_power,
    "from decimal import Decimal\nprint(Decimal('2') ** 3)\n",
    "8"
);
crate::runtime_case!(
    fractions_abs,
    "from fractions import Fraction\nprint(abs(Fraction(-3, 4)))\n",
    "3/4"
);
crate::runtime_case!(
    cmath_tanh,
    "import cmath\nprint(cmath.tanh(0))\n",
    "0j"
);
crate::runtime_case!(
    numbers_integral,
    "import numbers\nprint(issubclass(int, numbers.Integral))\n",
    "True"
);
crate::runtime_case!(
    numbers_rational,
    "from fractions import Fraction\nimport numbers\nprint(isinstance(Fraction(1, 2), numbers.Rational))\n",
    "True"
);
crate::runtime_case!(
    decimal_copy,
    "from decimal import Decimal\nprint(+Decimal('1.23'))\n",
    "1.23"
);
crate::runtime_case!(
    fractions_reduced,
    "from fractions import Fraction\nprint(Fraction(4, 8))\n",
    "1/2"
);
crate::runtime_case!(
    cmath_acos,
    "import cmath\nprint(cmath.acos(1))\n",
    "0j"
);

crate::compile_case!(decimal_localcontext, "from decimal import localcontext, Decimal\nwith localcontext() as ctx:\n Decimal('1')\n");
crate::compile_case!(decimal_clamp, "from decimal import Decimal, Context, Clamped\n");
crate::compile_case!(fractions_from_decimal, "from fractions import Fraction\nfrom decimal import Decimal\nFraction(Decimal('0.5'))\n");
crate::compile_case!(cmath_matrix, "import cmath\ncmath.sqrt(-1) * cmath.sqrt(-1)\n");
crate::compile_case!(decimal_compare_total, "from decimal import Decimal\nDecimal('1.0').compare_total(Decimal('1'))\n");

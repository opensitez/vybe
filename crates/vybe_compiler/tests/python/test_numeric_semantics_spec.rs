use super::helpers::*;

macro_rules! runtime_case {
    ($name:ident, $src:expr, $expected:expr) => {
        #[test]
        fn $name() {
            assert_eq!(run_python_one($src), $expected);
        }
    };
}

macro_rules! compile_case {
    ($name:ident, $src:expr) => {
        #[test]
        fn $name() {
            compile_ok($src);
        }
    };
}

runtime_case!(bin_literal_runtime, "print(0b101101)\n", "45");
runtime_case!(oct_literal_runtime, "print(0o17)\n", "15");
runtime_case!(hex_literal_runtime, "print(0x2A)\n", "42");
runtime_case!(int_underscore_runtime, "print(1_234_567)\n", "1234567");
compile_case!(float_underscore_compile, "x = 3.14_15_93\n");
runtime_case!(scientific_upper_runtime, "print(1E3)\n", "1000");
runtime_case!(scientific_negative_runtime, "print(1e-3)\n", "0.001");
compile_case!(complex_imaginary_compile, "z = 2j\n");
compile_case!(complex_add_compile, "z = (1 + 2j) + (3 + 4j)\n");
compile_case!(complex_sub_compile, "z = (4 + 3j) - (1 + 1j)\n");
runtime_case!(int_float_add_runtime, "print(1 + 2.5)\n", "3.5");
runtime_case!(int_float_mul_runtime, "print(2 * 2.5)\n", "5");
runtime_case!(int_float_div_runtime, "print(5 / 2)\n", "2.5");
runtime_case!(floor_div_negative_runtime, "print(-7 // 2)\n", "-4");
runtime_case!(modulo_negative_runtime, "print(-7 % 2)\n", "1");
runtime_case!(pow_negative_exponent_runtime, "print(2 ** -1)\n", "0.5");
runtime_case!(divmod_negative_runtime, "print(divmod(-7, 3)[0])\n", "-3");
runtime_case!(
    bool_is_int_runtime,
    "print(isinstance(True, int))\n",
    "true"
);
runtime_case!(int_base_two_runtime, "print(int('1010', 2))\n", "10");
runtime_case!(int_base_sixteen_runtime, "print(int('FF', 16))\n", "255");
compile_case!(float_from_scientific_compile, "x = float('1e6')\n");
compile_case!(complex_constructor_compile, "z = complex(3, 4)\n");
compile_case!(
    complex_real_imag_compile,
    "z = 3 + 4j\na = z.real\nb = z.imag\n"
);
compile_case!(unary_plus_literal_compile, "x = +42\n");
compile_case!(unary_minus_hex_compile, "x = -0x10\n");
runtime_case!(bitshift_large_literal_runtime, "print(1 << 10)\n", "1024");
compile_case!(round_negative_digits_compile, "x = round(1234, -2)\n");
runtime_case!(numeric_chain_mixed_runtime, "print(1 < 1.5 < 2)\n", "true");
compile_case!(inf_compare_compile, "x = float('inf') > 10\n");
compile_case!(nan_compare_compile, "x = float('nan') != float('nan')\n");

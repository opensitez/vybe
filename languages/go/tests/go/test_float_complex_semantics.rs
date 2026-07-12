//! float64 and complex128 semantics: math.IsNaN, IsInf, Inf, NaN; complex/real/imag;
//! cmplx.Abs, cmplx.Pow; Sprintf %e/%g; negative zero sign — distinct from
//! `test_math_extended.rs` and smoke in `test_stdlib_math_database.rs`.

go_run_cases! {
    math_isnan_on_nan => (
        "package main; import \"fmt\"; import \"math\"; func main() { fmt.Println(math.IsNaN(math.NaN())) }",
        vec!["true"]
    ),
    math_isnan_on_number => (
        "package main; import \"fmt\"; import \"math\"; func main() { fmt.Println(math.IsNaN(3.14)) }",
        vec!["false"]
    ),
    math_isnan_on_inf => (
        "package main; import \"fmt\"; import \"math\"; func main() { fmt.Println(math.IsNaN(math.Inf(1))) }",
        vec!["false"]
    ),
    math_isinf_positive => (
        "package main; import \"fmt\"; import \"math\"; func main() { fmt.Println(math.IsInf(math.Inf(1), 1)) }",
        vec!["true"]
    ),
    math_isinf_negative => (
        "package main; import \"fmt\"; import \"math\"; func main() { fmt.Println(math.IsInf(math.Inf(-1), -1)) }",
        vec!["true"]
    ),
    math_isinf_wrong_sign => (
        "package main; import \"fmt\"; import \"math\"; func main() { fmt.Println(math.IsInf(math.Inf(1), -1)) }",
        vec!["false"]
    ),
    math_isinf_on_finite => (
        "package main; import \"fmt\"; import \"math\"; func main() { fmt.Println(math.IsInf(1.0, 0)) }",
        vec!["false"]
    ),
    math_inf_positive_sign => (
        "package main; import \"fmt\"; import \"math\"; func main() { fmt.Println(math.IsInf(math.Inf(1), 0)) }",
        vec!["true"]
    ),
    math_inf_negative_sign => (
        "package main; import \"fmt\"; import \"math\"; func main() { fmt.Println(math.IsInf(math.Inf(-1), 0)) }",
        vec!["true"]
    ),
    math_nan_not_equal_self => (
        "package main; import \"fmt\"; import \"math\"; func main() { n := math.NaN(); fmt.Println(n == n) }",
        vec!["false"]
    ),
    math_nan_isnan_true => (
        "package main; import \"fmt\"; import \"math\"; func main() { fmt.Println(math.IsNaN(0.0/0.0)) }",
        vec!["true"]
    ),
    complex_literal_real_imag => (
        "package main; import \"fmt\"; func main() { z := complex(3, 4); fmt.Println(real(z)); fmt.Println(imag(z)) }",
        vec!["3", "4"]
    ),
    complex_zero => (
        "package main; import \"fmt\"; func main() { z := complex(0, 0); fmt.Println(real(z)); fmt.Println(imag(z)) }",
        vec!["0", "0"]
    ),
    complex_pure_imaginary => (
        "package main; import \"fmt\"; func main() { z := complex(0, 5); fmt.Println(real(z)); fmt.Println(imag(z)) }",
        vec!["0", "5"]
    ),
    complex_pure_real => (
        "package main; import \"fmt\"; func main() { z := complex(7, 0); fmt.Println(real(z)); fmt.Println(imag(z)) }",
        vec!["7", "0"]
    ),
    cmplx_abs_three_four => (
        "package main; import \"fmt\"; import \"math/cmplx\"; func main() { fmt.Println(cmplx.Abs(3 + 4i)) }",
        vec!["5"]
    ),
    cmplx_abs_zero => (
        "package main; import \"fmt\"; import \"math/cmplx\"; func main() { fmt.Println(cmplx.Abs(0)) }",
        vec!["0"]
    ),
    cmplx_abs_pure_imaginary => (
        "package main; import \"fmt\"; import \"math/cmplx\"; func main() { fmt.Println(cmplx.Abs(3i)) }",
        vec!["3"]
    ),
    cmplx_abs_negative_real => (
        "package main; import \"fmt\"; import \"math/cmplx\"; func main() { fmt.Println(cmplx.Abs(-3)) }",
        vec!["3"]
    ),
    cmplx_pow_square => (
        "package main; import \"fmt\"; import \"math/cmplx\"; func main() { z := cmplx.Pow(2, 2); fmt.Println(real(z)); fmt.Println(imag(z)) }",
        vec!["4", "0"]
    ),
    cmplx_pow_imaginary_unit => (
        "package main; import \"fmt\"; import \"math/cmplx\"; func main() { z := cmplx.Pow(1i, 2); fmt.Println(real(z)); fmt.Println(imag(z)) }",
        vec!["-1", "0"]
    ),
    cmplx_sqrt_negative_one => (
        "package main; import \"fmt\"; import \"math/cmplx\"; func main() { z := cmplx.Sqrt(-1); fmt.Println(real(z)); fmt.Println(imag(z) > 0) }",
        vec!["0", "true"]
    ),
    cmplx_conjugate => (
        "package main; import \"fmt\"; import \"math/cmplx\"; func main() { z := cmplx.Conj(3 + 4i); fmt.Println(real(z)); fmt.Println(imag(z)) }",
        vec!["3", "-4"]
    ),
    cmplx_exp_zero => (
        "package main; import \"fmt\"; import \"math/cmplx\"; func main() { z := cmplx.Exp(0); fmt.Println(real(z)); fmt.Println(imag(z)) }",
        vec!["1", "0"]
    ),
    cmplx_log_one => (
        "package main; import \"fmt\"; import \"math/cmplx\"; func main() { z := cmplx.Log(1); fmt.Println(real(z)); fmt.Println(imag(z)) }",
        vec!["0", "0"]
    ),
    cmplx_phase_positive_real => (
        "package main; import \"fmt\"; import \"math/cmplx\"; func main() { fmt.Println(cmplx.Phase(1)) }",
        vec!["0"]
    ),
    cmplx_polar_roundtrip => (
        "package main; import \"fmt\"; import \"math/cmplx\"; func main() { r, theta := cmplx.Polar(1); z := cmplx.Rect(r, theta); fmt.Println(real(z)); fmt.Println(imag(z)) }",
        vec!["1", "0"]
    ),
    fmt_sprintf_float_e_positive => (
        "package main; import \"fmt\"; func main() { fmt.Println(fmt.Sprintf(\"%.2e\", 1234.5)) }",
        vec!["1.23e+03"]
    ),
    fmt_sprintf_float_e_negative => (
        "package main; import \"fmt\"; import \"math\"; func main() { fmt.Println(fmt.Sprintf(\"%.2e\", -0.0012)) }",
        vec!["-1.20e-03"]
    ),
    fmt_sprintf_float_g_large => (
        "package main; import \"fmt\"; func main() { fmt.Println(fmt.Sprintf(\"%g\", 12345.0)) }",
        vec!["12345"]
    ),
    fmt_sprintf_float_g_small => (
        "package main; import \"fmt\"; func main() { fmt.Println(fmt.Sprintf(\"%g\", 0.0001234)) }",
        vec!["0.0001234"]
    ),
    fmt_sprintf_float_g_trailing_zeros_trimmed => (
        "package main; import \"fmt\"; func main() { fmt.Println(fmt.Sprintf(\"%g\", 1.500000)) }",
        vec!["1.5"]
    ),
    fmt_sprintf_float_f_fixed => (
        "package main; import \"fmt\"; func main() { fmt.Println(fmt.Sprintf(\"%.3f\", 1.23456)) }",
        vec!["1.235"]
    ),
    negative_zero_signbit => (
        "package main; import \"fmt\"; import \"math\"; func main() { z := math.Copysign(0, -1); fmt.Println(math.Signbit(z)) }",
        vec!["true"]
    ),
    positive_zero_not_signbit => (
        "package main; import \"fmt\"; import \"math\"; func main() { z := 0.0; fmt.Println(math.Signbit(z)) }",
        vec!["false"]
    ),
    negative_zero_equals_positive_zero => (
        "package main; import \"fmt\"; import \"math\"; func main() { z := math.Copysign(0, -1); fmt.Println(z == 0.0) }",
        vec!["true"]
    ),
    negative_zero_division => (
        "package main; import \"fmt\"; import \"math\"; func main() { z := math.Copysign(0, -1); fmt.Println(1.0/z < 0) }",
        vec!["true"]
    ),
    math_copysign_preserves_magnitude => (
        "package main; import \"fmt\"; import \"math\"; func main() { fmt.Println(math.Copysign(5, -1)) }",
        vec!["-5"]
    ),
    math_signbit_on_negative => (
        "package main; import \"fmt\"; import \"math\"; func main() { fmt.Println(math.Signbit(-2.5)) }",
        vec!["true"]
    ),
    math_signbit_on_positive => (
        "package main; import \"fmt\"; import \"math\"; func main() { fmt.Println(math.Signbit(2.5)) }",
        vec!["false"]
    ),
    complex_addition => (
        "package main; import \"fmt\"; func main() { z := (1 + 2i) + (3 + 4i); fmt.Println(real(z)); fmt.Println(imag(z)) }",
        vec!["4", "6"]
    ),
    complex_subtraction => (
        "package main; import \"fmt\"; func main() { z := (5 + 6i) - (2 + 1i); fmt.Println(real(z)); fmt.Println(imag(z)) }",
        vec!["3", "5"]
    ),
    complex_multiplication => (
        "package main; import \"fmt\"; func main() { z := (1 + 2i) * (3 + 4i); fmt.Println(real(z)); fmt.Println(imag(z)) }",
        vec!["-5", "10"]
    ),
    complex_division => (
        "package main; import \"fmt\"; func main() { z := (1 + 2i) / (1 + 1i); fmt.Println(real(z)); fmt.Println(imag(z)) }",
        vec!["1.5", "0.5"]
    ),
    cmplx_sin_zero => (
        "package main; import \"fmt\"; import \"math/cmplx\"; func main() { z := cmplx.Sin(0); fmt.Println(real(z)); fmt.Println(imag(z)) }",
        vec!["0", "0"]
    ),
    cmplx_cos_zero => (
        "package main; import \"fmt\"; import \"math/cmplx\"; func main() { z := cmplx.Cos(0); fmt.Println(real(z)); fmt.Println(imag(z)) }",
        vec!["1", "0"]
    ),
    cmplx_isnan_on_nan => (
        "package main; import \"fmt\"; import \"math/cmplx\"; import \"math\"; func main() { fmt.Println(cmplx.IsNaN(complex(math.NaN(), 0))) }",
        vec!["true"]
    ),
    cmplx_isinf_on_infinite => (
        "package main; import \"fmt\"; import \"math/cmplx\"; import \"math\"; func main() { fmt.Println(cmplx.IsInf(complex(math.Inf(1), 0))) }",
        vec!["true"]
    ),
    float64_equality_within_exact => (
        "package main; import \"fmt\"; func main() { fmt.Println(0.1+0.2 == 0.1+0.2) }",
        vec!["true"]
    ),
    math_float64bits_roundtrip => (
        "package main; import \"fmt\"; import \"math\"; func main() { bits := math.Float64bits(1.0); fmt.Println(math.Float64frombits(bits)) }",
        vec!["1"]
    ),
    math_float32bits_roundtrip => (
        "package main; import \"fmt\"; import \"math\"; func main() { bits := math.Float32bits(2.5); fmt.Println(math.Float32frombits(bits)) }",
        vec!["2.5"]
    ),
    fmt_sprintf_complex_default => (
        "package main; import \"fmt\"; func main() { fmt.Println(fmt.Sprintf(\"%v\", 1+2i)) }",
        vec!["(1+2i)"]
    ),
    fmt_sprintf_float_precision_zero => (
        "package main; import \"fmt\"; func main() { fmt.Println(fmt.Sprintf(\"%.0f\", 3.7)) }",
        vec!["4"]
    ),
    math_abs_float64 => (
        "package main; import \"fmt\"; import \"math\"; func main() { fmt.Println(math.Abs(-9.5)) }",
        vec!["9.5"]
    ),
    math_dim_float64 => (
        "package main; import \"fmt\"; import \"math\"; func main() { fmt.Println(math.Dim(3, 5)) }",
        vec!["0"]
    ),
    math_max_float64 => (
        "package main; import \"fmt\"; import \"math\"; func main() { fmt.Println(math.Max(3.0, 5.0)) }",
        vec!["5"]
    ),
    math_min_float64 => (
        "package main; import \"fmt\"; import \"math\"; func main() { fmt.Println(math.Min(3.0, 5.0)) }",
        vec!["3"]
    ),
}

go_compile_cases! {
    cmplx_pow_general_compile => "package main; import \"math/cmplx\"; func main() { _ = cmplx.Pow(1+1i, 2+0i) }",
    cmplx_tan_compile => "package main; import \"math/cmplx\"; func main() { _ = cmplx.Tan(1i) }",
    cmplx_asin_compile => "package main; import \"math/cmplx\"; func main() { _ = cmplx.Asin(0.5) }",
    cmplx_acos_compile => "package main; import \"math/cmplx\"; func main() { _ = cmplx.Acos(0.5) }",
    cmplx_atan_compile => "package main; import \"math/cmplx\"; func main() { _ = cmplx.Atan(1i) }",
    cmplx_sinh_compile => "package main; import \"math/cmplx\"; func main() { _ = cmplx.Sinh(1) }",
    cmplx_cosh_compile => "package main; import \"math/cmplx\"; func main() { _ = cmplx.Cosh(1) }",
    cmplx_tanh_compile => "package main; import \"math/cmplx\"; func main() { _ = cmplx.Tanh(1) }",
    math_inf_zero_compile => "package main; import \"math\"; func main() { _ = math.Inf(0) }",
    math_copysign_nan_compile => "package main; import \"math\"; func main() { _ = math.Copysign(1, math.NaN()) }",
    fmt_sprintf_float_g_auto_precision => "package main; import \"fmt\"; func main() { _ = fmt.Sprintf(\"%#g\", 3.14) }",
    complex128_type_alias => "package main; func main() { var z complex128 = complex(1, 2); _ = z }",
    float64_comparison_with_inf => "package main; import \"math\"; func main() { _ = 1.0 < math.Inf(1) }",
}

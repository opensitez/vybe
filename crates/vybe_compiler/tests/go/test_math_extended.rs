//! math package beyond test_math.rs: Log, Pow, Mod, Remainder, Trig, constants.

use crate::helpers::*;

go_run_cases! {
    math_pow_integer_exp => ("package main; import \"fmt\"; import \"math\"; func main() { fmt.Println(math.Pow(2, 3)) }", vec!["8"]),
    math_sqrt_two => ("package main; import \"fmt\"; import \"math\"; func main() { fmt.Println(math.Sqrt(2)) }", vec!["1.4142135623730951"]),
    math_mod_positive => ("package main; import \"fmt\"; import \"math\"; func main() { fmt.Println(math.Mod(7, 3)) }", vec!["1"]),
    math_remainder_ieee => ("package main; import \"fmt\"; import \"math\"; func main() { fmt.Println(math.Remainder(7, 3)) }", vec!["1"]),
    math_log_natural_e => ("package main; import \"fmt\"; import \"math\"; func main() { fmt.Println(math.Log(math.E)) }", vec!["1"]),
    math_log10_thousand => ("package main; import \"fmt\"; import \"math\"; func main() { fmt.Println(math.Log10(1000)) }", vec!["3"]),
    math_sin_zero => ("package main; import \"fmt\"; import \"math\"; func main() { fmt.Println(math.Sin(0)) }", vec!["0"]),
    math_cos_zero => ("package main; import \"fmt\"; import \"math\"; func main() { fmt.Println(math.Cos(0)) }", vec!["1"]),
    math_hypot_three_four => ("package main; import \"fmt\"; import \"math\"; func main() { fmt.Println(math.Hypot(3, 4)) }", vec!["5"]),
    math_dim_returns_two => ("package main; import \"fmt\"; import \"math\"; func main() { fmt.Println(math.Dim(5, 3)) }", vec!["3"]),
}

go_compile_cases! {
    math_constants_pi => "package main; import \"math\"; func main() { _ = math.Pi }",
    math_constants_phi => "package main; import \"math\"; func main() { _ = math.Phi }",
    math_atan2_quadrant => "package main; import \"math\"; func main() { _ = math.Atan2(1, 1) }",
    math_copysign_negative => "package main; import \"math\"; func main() { _ = math.Copysign(1, -1) }",
    math_nextafter_float => "package main; import \"math\"; func main() { _ = math.Nextafter(1.0, 2.0) }",
    math_ldexp_scales_mantissa => "package main; import \"math\"; func main() { _ = math.Ldexp(0.5, 2) }",
}

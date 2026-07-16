//! math/big runtime: Int SetString, Add, Sub, Mul, Div, Mod, Cmp, String bases,
//! Bit, GCD, ProbablyPrime, Rat Add, Float64 — distinct from smoke in
//! `test_stdlib_math_database.rs`.

go_run_cases! {
    big_int_set_string_base10 => (
        "package main; import \"fmt\"; import \"math/big\"; func main() { z := new(big.Int); z.SetString(\"12345\", 10); fmt.Println(z.String()) }",
        vec!["12345"]
    ),
    big_int_set_string_base16 => (
        "package main; import \"fmt\"; import \"math/big\"; func main() { z := new(big.Int); z.SetString(\"ff\", 16); fmt.Println(z.String()) }",
        vec!["255"]
    ),
    big_int_set_string_base2 => (
        "package main; import \"fmt\"; import \"math/big\"; func main() { z := new(big.Int); z.SetString(\"1010\", 2); fmt.Println(z.String()) }",
        vec!["10"]
    ),
    big_int_set_string_invalid_returns_false => (
        "package main; import \"fmt\"; import \"math/big\"; func main() { z := new(big.Int); _, ok := z.SetString(\"12z\", 10); fmt.Println(ok) }",
        vec!["false"]
    ),
    big_int_add_positive => (
        "package main; import \"fmt\"; import \"math/big\"; func main() { a := big.NewInt(100); b := big.NewInt(23); fmt.Println(a.Add(a, b).String()) }",
        vec!["123"]
    ),
    big_int_add_negative => (
        "package main; import \"fmt\"; import \"math/big\"; func main() { a := big.NewInt(-5); b := big.NewInt(3); fmt.Println(a.Add(a, b).String()) }",
        vec!["-2"]
    ),
    big_int_sub_basic => (
        "package main; import \"fmt\"; import \"math/big\"; func main() { a := big.NewInt(10); b := big.NewInt(4); fmt.Println(a.Sub(a, b).String()) }",
        vec!["6"]
    ),
    big_int_sub_negative_result => (
        "package main; import \"fmt\"; import \"math/big\"; func main() { a := big.NewInt(3); b := big.NewInt(10); fmt.Println(a.Sub(a, b).String()) }",
        vec!["-7"]
    ),
    big_int_mul_small => (
        "package main; import \"fmt\"; import \"math/big\"; func main() { a := big.NewInt(6); b := big.NewInt(7); fmt.Println(a.Mul(a, b).String()) }",
        vec!["42"]
    ),
    big_int_mul_by_zero => (
        "package main; import \"fmt\"; import \"math/big\"; func main() { a := big.NewInt(99); b := big.NewInt(0); fmt.Println(a.Mul(a, b).String()) }",
        vec!["0"]
    ),
    big_int_div_exact => (
        "package main; import \"fmt\"; import \"math/big\"; func main() { a := big.NewInt(20); b := big.NewInt(4); fmt.Println(a.Div(a, b).String()) }",
        vec!["5"]
    ),
    big_int_div_truncates => (
        "package main; import \"fmt\"; import \"math/big\"; func main() { a := big.NewInt(7); b := big.NewInt(2); fmt.Println(a.Div(a, b).String()) }",
        vec!["3"]
    ),
    big_int_mod_positive => (
        "package main; import \"fmt\"; import \"math/big\"; func main() { a := big.NewInt(17); b := big.NewInt(5); fmt.Println(a.Mod(a, b).String()) }",
        vec!["2"]
    ),
    big_int_mod_by_one => (
        "package main; import \"fmt\"; import \"math/big\"; func main() { a := big.NewInt(42); b := big.NewInt(1); fmt.Println(a.Mod(a, b).String()) }",
        vec!["0"]
    ),
    big_int_cmp_equal => (
        "package main; import \"fmt\"; import \"math/big\"; func main() { a := big.NewInt(5); b := big.NewInt(5); fmt.Println(a.Cmp(b)) }",
        vec!["0"]
    ),
    big_int_cmp_less => (
        "package main; import \"fmt\"; import \"math/big\"; func main() { a := big.NewInt(3); b := big.NewInt(7); fmt.Println(a.Cmp(b)) }",
        vec!["-1"]
    ),
    big_int_cmp_greater => (
        "package main; import \"fmt\"; import \"math/big\"; func main() { a := big.NewInt(9); b := big.NewInt(2); fmt.Println(a.Cmp(b)) }",
        vec!["1"]
    ),
    big_int_string_base10 => (
        "package main; import \"fmt\"; import \"math/big\"; func main() { z := big.NewInt(9876); fmt.Println(z.String()) }",
        vec!["9876"]
    ),
    big_int_format_base16 => (
        "package main; import \"fmt\"; import \"math/big\"; func main() { z := big.NewInt(255); fmt.Println(z.Text(16)) }",
        vec!["ff"]
    ),
    big_int_format_base8 => (
        "package main; import \"fmt\"; import \"math/big\"; func main() { z := big.NewInt(8); fmt.Println(z.Text(8)) }",
        vec!["10"]
    ),
    big_int_bit_and => (
        "package main; import \"fmt\"; import \"math/big\"; func main() { a := big.NewInt(12); b := big.NewInt(10); fmt.Println(a.And(a, b).String()) }",
        vec!["8"]
    ),
    big_int_bit_or => (
        "package main; import \"fmt\"; import \"math/big\"; func main() { a := big.NewInt(5); b := big.NewInt(3); fmt.Println(a.Or(a, b).String()) }",
        vec!["7"]
    ),
    big_int_bit_xor => (
        "package main; import \"fmt\"; import \"math/big\"; func main() { a := big.NewInt(15); b := big.NewInt(10); fmt.Println(a.Xor(a, b).String()) }",
        vec!["5"]
    ),
    big_int_bit_not => (
        "package main; import \"fmt\"; import \"math/big\"; func main() { a := big.NewInt(0); fmt.Println(a.Not(a).String()) }",
        vec!["-1"]
    ),
    big_int_bit_len => (
        "package main; import \"fmt\"; import \"math/big\"; func main() { z := big.NewInt(8); fmt.Println(z.BitLen()) }",
        vec!["4"]
    ),
    big_int_bit_set_and_test => (
        "package main; import \"fmt\"; import \"math/big\"; func main() { z := big.NewInt(0); z.SetBit(z, 3, 1); fmt.Println(z.Bit(3)) }",
        vec!["1"]
    ),
    big_int_gcd => (
        "package main; import \"fmt\"; import \"math/big\"; func main() { a := big.NewInt(48); b := big.NewInt(18); fmt.Println(new(big.Int).GCD(nil, nil, a, b).String()) }",
        vec!["6"]
    ),
    big_int_gcd_coprime => (
        "package main; import \"fmt\"; import \"math/big\"; func main() { a := big.NewInt(17); b := big.NewInt(13); fmt.Println(new(big.Int).GCD(nil, nil, a, b).String()) }",
        vec!["1"]
    ),
    big_int_abs_negative => (
        "package main; import \"fmt\"; import \"math/big\"; func main() { z := big.NewInt(-42); fmt.Println(z.Abs(z).String()) }",
        vec!["42"]
    ),
    big_int_neg => (
        "package main; import \"fmt\"; import \"math/big\"; func main() { z := big.NewInt(7); fmt.Println(z.Neg(z).String()) }",
        vec!["-7"]
    ),
    big_int_sign_positive => (
        "package main; import \"fmt\"; import \"math/big\"; func main() { z := big.NewInt(5); fmt.Println(z.Sign()) }",
        vec!["1"]
    ),
    big_int_sign_zero => (
        "package main; import \"fmt\"; import \"math/big\"; func main() { z := big.NewInt(0); fmt.Println(z.Sign()) }",
        vec!["0"]
    ),
    big_int_sign_negative => (
        "package main; import \"fmt\"; import \"math/big\"; func main() { z := big.NewInt(-3); fmt.Println(z.Sign()) }",
        vec!["-1"]
    ),
    big_int_lsh => (
        "package main; import \"fmt\"; import \"math/big\"; func main() { z := big.NewInt(1); fmt.Println(z.Lsh(z, 4).String()) }",
        vec!["16"]
    ),
    big_int_rsh => (
        "package main; import \"fmt\"; import \"math/big\"; func main() { z := big.NewInt(32); fmt.Println(z.Rsh(z, 3).String()) }",
        vec!["4"]
    ),
    big_int_exp_small => (
        "package main; import \"fmt\"; import \"math/big\"; func main() { z := big.NewInt(2); fmt.Println(z.Exp(z, big.NewInt(10), nil).String()) }",
        vec!["1024"]
    ),
    big_int_is_probably_prime_small => (
        "package main; import \"fmt\"; import \"math/big\"; func main() { z := big.NewInt(17); fmt.Println(z.ProbablyPrime(10)) }",
        vec!["true"]
    ),
    big_int_is_probably_prime_composite => (
        "package main; import \"fmt\"; import \"math/big\"; func main() { z := big.NewInt(15); fmt.Println(z.ProbablyPrime(10)) }",
        vec!["false"]
    ),
    big_rat_add => (
        "package main; import \"fmt\"; import \"math/big\"; func main() { a := big.NewRat(1, 3); b := big.NewRat(1, 6); fmt.Println(a.Add(a, b).FloatString(2)) }",
        vec!["0.50"]
    ),
    big_rat_sub => (
        "package main; import \"fmt\"; import \"math/big\"; func main() { a := big.NewRat(3, 4); b := big.NewRat(1, 4); fmt.Println(a.Sub(a, b).FloatString(2)) }",
        vec!["0.50"]
    ),
    big_rat_mul => (
        "package main; import \"fmt\"; import \"math/big\"; func main() { a := big.NewRat(2, 3); b := big.NewRat(3, 2); fmt.Println(a.Mul(a, b).FloatString(2)) }",
        vec!["1.00"]
    ),
    big_rat_float64 => (
        "package main; import \"fmt\"; import \"math/big\"; func main() { r := big.NewRat(1, 4); f, _ := r.Float64(); fmt.Println(f) }",
        vec!["0.25"]
    ),
    big_rat_string => (
        "package main; import \"fmt\"; import \"math/big\"; func main() { r := big.NewRat(22, 7); fmt.Println(r.String()) }",
        vec!["22/7"]
    ),
    big_float_add => (
        "package main; import \"fmt\"; import \"math/big\"; func main() { a := big.NewFloat(1.5); b := big.NewFloat(2.5); fmt.Println(a.Add(a, b).String()) }",
        vec!["4"]
    ),
    big_float_sub => (
        "package main; import \"fmt\"; import \"math/big\"; func main() { a := big.NewFloat(5.0); b := big.NewFloat(2.0); fmt.Println(a.Sub(a, b).String()) }",
        vec!["3"]
    ),
    big_float_mul => (
        "package main; import \"fmt\"; import \"math/big\"; func main() { a := big.NewFloat(2.0); b := big.NewFloat(3.0); fmt.Println(a.Mul(a, b).String()) }",
        vec!["6"]
    ),
    big_float_float64 => (
        "package main; import \"fmt\"; import \"math/big\"; func main() { f := big.NewFloat(3.14); v, _ := f.Float64(); fmt.Println(v) }",
        vec!["3.14"]
    ),
    big_int_new_int_zero => (
        "package main; import \"fmt\"; import \"math/big\"; func main() { z := big.NewInt(0); fmt.Println(z.String()) }",
        vec!["0"]
    ),
    big_int_quo_rem => (
        "package main; import \"fmt\"; import \"math/big\"; func main() { a := big.NewInt(23); b := big.NewInt(5); q := new(big.Int); r := new(big.Int); q.QuoRem(a, b, r); fmt.Println(q.String()); fmt.Println(r.String()) }",
        vec!["4", "3"]
    ),
    big_int_bytes_roundtrip => (
        "package main; import \"fmt\"; import \"math/big\"; func main() { orig := big.NewInt(1000); back := new(big.Int).SetBytes(orig.Bytes()); fmt.Println(back.String()) }",
        vec!["1000"]
    ),
}

go_compile_cases! {
    big_int_probably_prime_large => "package main; import \"math/big\"; func main() { z := new(big.Int); z.SetString(\"982451653\", 10); _ = z.ProbablyPrime(20) }",
    big_rat_inv => "package main; import \"math/big\"; func main() { r := big.NewRat(2, 3); _ = r.Inv(r) }",
    big_rat_set_frac => "package main; import \"math/big\"; func main() { r := new(big.Rat); r.SetFrac(big.NewInt(3), big.NewInt(4)) }",
    big_rat_set_int64 => "package main; import \"math/big\"; func main() { r := new(big.Rat); r.SetInt64(7) }",
    big_rat_set_float64 => "package main; import \"math/big\"; func main() { r := new(big.Rat); _ = r.SetFloat64(0.75) }",
    big_float_parse => "package main; import \"math/big\"; func main() { f := new(big.Float); _, _ = f.SetString(\"1.23e2\") }",
    big_float_set_int => "package main; import \"math/big\"; func main() { f := new(big.Float); f.SetInt(big.NewInt(42)) }",
    big_int_rand_bits => "package main; import \"math/big\"; func main() { _ = big.NewInt(0).SetBit(nil, 64, 1) }",
    big_int_mod_inverse => "package main; import \"math/big\"; func main() { a := big.NewInt(3); m := big.NewInt(11); _ = new(big.Int).ModInverse(a, m) }",
    big_int_sqrt => "package main; import \"math/big\"; func main() { z := big.NewInt(16); _ = new(big.Int).Sqrt(z) }",
}

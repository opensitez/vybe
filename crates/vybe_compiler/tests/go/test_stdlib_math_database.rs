//! math/big and math/cmplx — one smoke per distinct API.


go_run_cases! {
    big_int_add => ("package main; import \"fmt\"; import \"math/big\"; func main() { a := big.NewInt(10); b := big.NewInt(7); fmt.Println(a.Add(a, b).String()) }", vec!["17"]),
    cmplx_sqrt_neg_one => ("package main; import \"fmt\"; import \"math/cmplx\"; func main() { z := cmplx.Sqrt(-1); fmt.Println(cmplx.Imag(z) > 0) }", vec!["true"]),
}

go_compile_cases! {
    big_new_rat => "package main; import \"math/big\"; func main() { _ = big.NewRat(1, 2) }",
    big_new_float => "package main; import \"math/big\"; func main() { _ = big.NewFloat(1.5) }",
    big_int_set_string => "package main; import \"math/big\"; func main() { z := new(big.Int); _, _ = z.SetString(\"ff\", 16) }",
    cmplx_abs => "package main; import \"math/cmplx\"; func main() { _ = cmplx.Abs(3 + 4i) }",
    cmplx_polar => "package main; import \"math/cmplx\"; func main() { _, _ = cmplx.Polar(1 + 1i) }",
}

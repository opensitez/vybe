// vybe-test: go/float_complex_semantics/cmplx_pow_general_compile
// origin: languages/go/tests/go/test_float_complex_semantics.rs
// vybe-test-mode: compile

package main
import "math/cmplx"
func main() { _ = cmplx.Pow(1+1i, 2+0i) }

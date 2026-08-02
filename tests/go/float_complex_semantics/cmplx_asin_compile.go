// vybe-test: go/float_complex_semantics/cmplx_asin_compile
// origin: languages/go/tests/go/test_float_complex_semantics.rs
// vybe-test-mode: compile

package main
import "math/cmplx"
func main() { _ = cmplx.Asin(0.5) }

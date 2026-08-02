// vybe-test: go/float_complex_semantics/math_inf_zero_compile
// origin: languages/go/tests/go/test_float_complex_semantics.rs
// vybe-test-mode: compile

package main
import "math"
func main() { _ = math.Inf(0) }

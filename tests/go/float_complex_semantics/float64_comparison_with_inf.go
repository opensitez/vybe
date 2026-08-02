// vybe-test: go/float_complex_semantics/float64_comparison_with_inf
// origin: languages/go/tests/go/test_float_complex_semantics.rs
// vybe-test-mode: compile

package main
import "math"
func main() { _ = 1.0 < math.Inf(1) }

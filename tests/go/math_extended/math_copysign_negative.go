// vybe-test: go/math_extended/math_copysign_negative
// origin: languages/go/tests/go/test_math_extended.rs
// vybe-test-mode: compile

package main
import "math"
func main() { _ = math.Copysign(1, -1) }

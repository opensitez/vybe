// vybe-test: go/math_big_int/big_int_sqrt
// origin: languages/go/tests/go/test_math_big_int.rs
// vybe-test-mode: compile

package main
import "math/big"
func main() { z := big.NewInt(16)
_ = new(big.Int).Sqrt(z) }

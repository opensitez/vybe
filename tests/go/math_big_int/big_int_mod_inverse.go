// vybe-test: go/math_big_int/big_int_mod_inverse
// origin: languages/go/tests/go/test_math_big_int.rs
// vybe-test-mode: compile

package main
import "math/big"
func main() { a := big.NewInt(3)
m := big.NewInt(11)
_ = new(big.Int).ModInverse(a, m) }

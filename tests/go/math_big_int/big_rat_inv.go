// vybe-test: go/math_big_int/big_rat_inv
// origin: languages/go/tests/go/test_math_big_int.rs
// vybe-test-mode: compile

package main
import "math/big"
func main() { r := big.NewRat(2, 3)
_ = r.Inv(r) }

// vybe-test: go/math_big_int/big_rat_set_frac
// origin: languages/go/tests/go/test_math_big_int.rs
// vybe-test-mode: compile

package main
import "math/big"
func main() { r := new(big.Rat)
r.SetFrac(big.NewInt(3), big.NewInt(4)) }

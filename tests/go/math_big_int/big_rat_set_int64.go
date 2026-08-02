// vybe-test: go/math_big_int/big_rat_set_int64
// origin: languages/go/tests/go/test_math_big_int.rs
// vybe-test-mode: compile

package main
import "math/big"
func main() { r := new(big.Rat)
r.SetInt64(7) }

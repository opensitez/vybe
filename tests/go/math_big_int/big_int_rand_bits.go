// vybe-test: go/math_big_int/big_int_rand_bits
// origin: languages/go/tests/go/test_math_big_int.rs
// vybe-test-mode: compile

package main
import "math/big"
func main() { _ = big.NewInt(0).SetBit(nil, 64, 1) }

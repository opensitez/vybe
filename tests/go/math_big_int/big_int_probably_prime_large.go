// vybe-test: go/math_big_int/big_int_probably_prime_large
// origin: languages/go/tests/go/test_math_big_int.rs
// vybe-test-mode: compile

package main
import "math/big"
func main() { z := new(big.Int)
z.SetString("982451653", 10)
_ = z.ProbablyPrime(20) }

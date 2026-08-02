// vybe-test: go/math_big_int/big_int_is_probably_prime_small
// origin: languages/go/tests/go/test_math_big_int.rs

package main
import "fmt"
import "math/big"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { z := big.NewInt(17)
__check(fmt.Sprint(z.ProbablyPrime(10)), "true") }

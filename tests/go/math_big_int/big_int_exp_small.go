// vybe-test: go/math_big_int/big_int_exp_small
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

func main() { z := big.NewInt(2)
__check(fmt.Sprint(z.Exp(z, big.NewInt(10), nil).String()), "1024") }

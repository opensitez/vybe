// vybe-test: go/math_big_int/big_int_neg
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

func main() { z := big.NewInt(7)
__check(fmt.Sprint(z.Neg(z).String()), "-7") }

// vybe-test: go/math_big_int/big_int_mul_by_zero
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

func main() { a := big.NewInt(99)
b := big.NewInt(0)
__check(fmt.Sprint(a.Mul(a, b).String()), "0") }

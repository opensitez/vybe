// vybe-test: go/math_big_int/big_int_bit_or
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

func main() { a := big.NewInt(5)
b := big.NewInt(3)
__check(fmt.Sprint(a.Or(a, b).String()), "7") }

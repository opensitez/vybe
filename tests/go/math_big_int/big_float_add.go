// vybe-test: go/math_big_int/big_float_add
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

func main() { a := big.NewFloat(1.5)
b := big.NewFloat(2.5)
__check(fmt.Sprint(a.Add(a, b).String()), "4") }

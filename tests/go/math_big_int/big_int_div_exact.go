// vybe-test: go/math_big_int/big_int_div_exact
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

func main() { a := big.NewInt(20)
b := big.NewInt(4)
__check(fmt.Sprint(a.Div(a, b).String()), "5") }

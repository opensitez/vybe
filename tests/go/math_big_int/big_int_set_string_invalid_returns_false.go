// vybe-test: go/math_big_int/big_int_set_string_invalid_returns_false
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

func main() { z := new(big.Int)
_, ok := z.SetString("12z", 10)
__check(fmt.Sprint(ok), "false") }

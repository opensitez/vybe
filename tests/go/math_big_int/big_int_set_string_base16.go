// vybe-test: go/math_big_int/big_int_set_string_base16
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
z.SetString("ff", 16)
__check(fmt.Sprint(z.String()), "255") }

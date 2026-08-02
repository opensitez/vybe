// vybe-test: go/math_big_int/big_int_rsh
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

func main() { z := big.NewInt(32)
__check(fmt.Sprint(z.Rsh(z, 3).String()), "4") }

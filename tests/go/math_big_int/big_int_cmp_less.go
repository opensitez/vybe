// vybe-test: go/math_big_int/big_int_cmp_less
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

func main() { a := big.NewInt(3)
b := big.NewInt(7)
__check(fmt.Sprint(a.Cmp(b)), "-1") }

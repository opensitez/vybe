// vybe-test: go/math_big_int/big_int_quo_rem
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

func main() { a := big.NewInt(23)
b := big.NewInt(5)
q := new(big.Int)
r := new(big.Int)
q.QuoRem(a, b, r)
__check(fmt.Sprint(q.String()), "4")
__check(fmt.Sprint(r.String()), "3") }

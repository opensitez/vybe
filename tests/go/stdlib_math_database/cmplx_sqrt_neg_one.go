// vybe-test: go/stdlib_math_database/cmplx_sqrt_neg_one
// origin: languages/go/tests/go/test_stdlib_math_database.rs

package main
import "fmt"
import "math/cmplx"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { z := cmplx.Sqrt(-1)
__check(fmt.Sprint(cmplx.Imag(z) > 0), "true") }

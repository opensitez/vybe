// vybe-test: go/float_complex_semantics/cmplx_pow_imaginary_unit
// origin: languages/go/tests/go/test_float_complex_semantics.rs

package main
import "fmt"
import "math/cmplx"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { z := cmplx.Pow(1i, 2)
__check(fmt.Sprint(real(z)), "-1")
__check(fmt.Sprint(imag(z)), "0") }

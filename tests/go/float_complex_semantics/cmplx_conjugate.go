// vybe-test: go/float_complex_semantics/cmplx_conjugate
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

func main() { z := cmplx.Conj(3 + 4i)
__check(fmt.Sprint(real(z)), "3")
__check(fmt.Sprint(imag(z)), "-4") }

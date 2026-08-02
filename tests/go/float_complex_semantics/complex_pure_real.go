// vybe-test: go/float_complex_semantics/complex_pure_real
// origin: languages/go/tests/go/test_float_complex_semantics.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { z := complex(7, 0)
__check(fmt.Sprint(real(z)), "7")
__check(fmt.Sprint(imag(z)), "0") }

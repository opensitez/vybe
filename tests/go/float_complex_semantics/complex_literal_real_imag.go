// vybe-test: go/float_complex_semantics/complex_literal_real_imag
// origin: languages/go/tests/go/test_float_complex_semantics.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { z := complex(3, 4)
__check(fmt.Sprint(real(z)), "3")
__check(fmt.Sprint(imag(z)), "4") }

// vybe-test: go/float_complex_semantics/complex_subtraction
// origin: languages/go/tests/go/test_float_complex_semantics.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { z := (5 + 6i) - (2 + 1i)
__check(fmt.Sprint(real(z)), "3")
__check(fmt.Sprint(imag(z)), "5") }

// vybe-test: go/float_complex_semantics/complex_division
// origin: languages/go/tests/go/test_float_complex_semantics.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { z := (1 + 2i) / (1 + 1i)
__check(fmt.Sprint(real(z)), "1.5")
__check(fmt.Sprint(imag(z)), "0.5") }

// vybe-test: go/float_complex_semantics/complex_addition
// origin: languages/go/tests/go/test_float_complex_semantics.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { z := (1 + 2i) + (3 + 4i)
__check(fmt.Sprint(real(z)), "4")
__check(fmt.Sprint(imag(z)), "6") }

// vybe-test: go/float_complex_semantics/cmplx_pow_imaginary_unit
// origin: languages/go/tests/go/test_float_complex_semantics.rs

package main
import "fmt"
import "math/cmplx"
var __buf string

// __p appends one line, __pr appends without a newline.
func __p(s string) { __buf = __buf + s + "\n" }

func __pr(s string) { __buf = __buf + s }

// __check ends the program unless the collected output equals want. The final
// Println contributes a trailing newline the expected line vector never
// carried, so both forms are accepted.
func __check(want string) {
	if __buf != want && __buf != want+"\n" {
		fmt.Println("FAIL: want [" + want + "] got [" + __buf + "]")
		panic("assertion failed")
	}
}

func main() { z := cmplx.Pow(1i, 2)
__p(fmt.Sprint(real(z)))
__p(fmt.Sprint(imag(z))) 
__check("-1\n0")
}

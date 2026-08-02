// vybe-test: go/float_complex_semantics/cmplx_isinf_on_infinite
// origin: languages/go/tests/go/test_float_complex_semantics.rs

package main
import "fmt"
import "math/cmplx"
import "math"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(cmplx.IsInf(complex(math.Inf(1), 0))), "true") }

// vybe-test: go/float_complex_semantics/math_nan_not_equal_self
// origin: languages/go/tests/go/test_float_complex_semantics.rs

package main
import "fmt"
import "math"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { n := math.NaN()
__check(fmt.Sprint(n == n), "false") }

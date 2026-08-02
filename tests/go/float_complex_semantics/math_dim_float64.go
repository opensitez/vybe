// vybe-test: go/float_complex_semantics/math_dim_float64
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

func main() { __check(fmt.Sprint(math.Dim(3, 5)), "0") }

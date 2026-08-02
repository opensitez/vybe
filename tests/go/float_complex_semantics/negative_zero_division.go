// vybe-test: go/float_complex_semantics/negative_zero_division
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

func main() { z := math.Copysign(0, -1)
__check(fmt.Sprint(1.0/z < 0), "true") }

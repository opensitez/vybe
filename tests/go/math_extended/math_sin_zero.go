// vybe-test: go/math_extended/math_sin_zero
// origin: languages/go/tests/go/test_math_extended.rs

package main
import "fmt"
import "math"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(math.Sin(0)), "0") }

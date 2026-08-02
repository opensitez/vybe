// vybe-test: go/float_complex_semantics/float64_equality_within_exact
// origin: languages/go/tests/go/test_float_complex_semantics.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(0.1+0.2 == 0.1+0.2), "true") }

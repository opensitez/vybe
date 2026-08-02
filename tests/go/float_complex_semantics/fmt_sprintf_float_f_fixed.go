// vybe-test: go/float_complex_semantics/fmt_sprintf_float_f_fixed
// origin: languages/go/tests/go/test_float_complex_semantics.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(fmt.Sprintf("%.3f", 1.23456)), "1.235") }

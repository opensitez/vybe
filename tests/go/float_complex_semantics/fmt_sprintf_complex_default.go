// vybe-test: go/float_complex_semantics/fmt_sprintf_complex_default
// origin: languages/go/tests/go/test_float_complex_semantics.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(fmt.Sprintf("%v", 1+2i)), "(1+2i)") }

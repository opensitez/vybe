// vybe-test: go/function_types_advanced/two_nil_funcs_different_signatures_both_nil
// origin: languages/go/tests/go/test_function_types_advanced.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var noop func()
var add func(int) int
__check(fmt.Sprint(noop == nil), "true")
__check(fmt.Sprint(add == nil), "true") }

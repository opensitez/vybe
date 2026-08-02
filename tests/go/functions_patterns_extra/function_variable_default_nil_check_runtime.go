// vybe-test: go/functions_patterns_extra/function_variable_default_nil_check_runtime
// origin: languages/go/tests/go/test_functions_patterns_extra.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var fn func(int) int
__check(fmt.Sprint(fn == nil), "true")
}

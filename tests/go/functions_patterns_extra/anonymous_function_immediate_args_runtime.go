// vybe-test: go/functions_patterns_extra/anonymous_function_immediate_args_runtime
// origin: languages/go/tests/go/test_functions_patterns_extra.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(func(a int, b int) int { return a - b }(9, 4)), "5")
}

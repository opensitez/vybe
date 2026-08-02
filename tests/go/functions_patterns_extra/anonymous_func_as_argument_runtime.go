// vybe-test: go/functions_patterns_extra/anonymous_func_as_argument_runtime
// origin: languages/go/tests/go/test_functions_patterns_extra.rs

package main
import "fmt"
func consume(fn func() string) string { return fn() }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(consume(func() string { return "ok" })), "ok")
}

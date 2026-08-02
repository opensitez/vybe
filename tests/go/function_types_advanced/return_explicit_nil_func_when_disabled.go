// vybe-test: go/function_types_advanced/return_explicit_nil_func_when_disabled
// origin: languages/go/tests/go/test_function_types_advanced.rs

package main
import "fmt"
func maybe(enabled bool) func() { if enabled { return func() {} }
return nil }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(maybe(false) == nil), "true")
__check(fmt.Sprint(maybe(true) == nil), "false") }

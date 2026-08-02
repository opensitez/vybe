// vybe-test: go/lang_functions_returns/function_as_value_nil_compare
// origin: languages/go/tests/go/test_lang_functions_returns.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var f func()
__check(fmt.Sprint(f == nil), "true") }

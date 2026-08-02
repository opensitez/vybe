// vybe-test: go/lang_functions_returns/package_level_func_forward_ref
// origin: languages/go/tests/go/test_lang_functions_returns.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(g()), "5") }
func g() int { return 5 }

// vybe-test: go/lang_functions_returns/defer_modifies_named_result
// origin: languages/go/tests/go/test_lang_functions_returns.rs

package main
import "fmt"
func f() (n int) { defer func() { n++ }()
return 1 }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(f()), "2") }

// vybe-test: go/lang_functions_returns/named_result_assignment_before_bare_return
// origin: languages/go/tests/go/test_lang_functions_returns.rs

package main
import "fmt"
func f() (n int) { n = 9
return }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(f()), "9") }

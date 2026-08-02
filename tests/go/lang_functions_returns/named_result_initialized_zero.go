// vybe-test: go/lang_functions_returns/named_result_initialized_zero
// origin: languages/go/tests/go/test_lang_functions_returns.rs

package main
import "fmt"
func f() (n int, s string) { return }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { n, s := f()
__check(fmt.Sprint(n) + " " + fmt.Sprint(s == ""), "0 true") }

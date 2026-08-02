// vybe-test: go/lang_functions_returns/multi_return_ignore_with_blank
// origin: languages/go/tests/go/test_lang_functions_returns.rs

package main
import "fmt"
func pair() (int, string) { return 2, "b" }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { _, s := pair()
__check(fmt.Sprint(s), "b") }

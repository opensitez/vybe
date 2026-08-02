// vybe-test: go/lang_functions_returns/multi_return_swap
// origin: languages/go/tests/go/test_lang_functions_returns.rs

package main
import "fmt"
func pair() (int, string) { return 1, "a" }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { a, b := pair()
__check(fmt.Sprint(a) + " " + fmt.Sprint(b), "1 a") }

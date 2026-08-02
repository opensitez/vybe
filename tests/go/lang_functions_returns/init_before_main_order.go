// vybe-test: go/lang_functions_returns/init_before_main_order
// origin: languages/go/tests/go/test_lang_functions_returns.rs

package main
import "fmt"
var n = func() int { return 3 }()
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(n), "3") }

// vybe-test: go/lang_builtins_control/defer_lifo_order
// origin: languages/go/tests/go/test_lang_builtins_control.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { defer __check(fmt.Sprint(1), "1")
defer __check(fmt.Sprint(2), "2")
__check(fmt.Sprint(0), "0") }

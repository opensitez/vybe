// vybe-test: go/lang_builtins_control/multiple_assignment_swap
// origin: languages/go/tests/go/test_lang_builtins_control.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { a, b := 1, 2
a, b = b, a
__check(fmt.Sprint(a) + " " + fmt.Sprint(b), "2 1") }

// vybe-test: go/lang_builtins_control/blank_identifier_discard
// origin: languages/go/tests/go/test_lang_builtins_control.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { _, y := 1, 2
__check(fmt.Sprint(y), "2") }

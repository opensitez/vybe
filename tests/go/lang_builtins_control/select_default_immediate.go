// vybe-test: go/lang_builtins_control/select_default_immediate
// origin: languages/go/tests/go/test_lang_builtins_control.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { select { default: __check(fmt.Sprint("d"), "d") } }

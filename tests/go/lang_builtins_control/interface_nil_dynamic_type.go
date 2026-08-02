// vybe-test: go/lang_builtins_control/interface_nil_dynamic_type
// origin: languages/go/tests/go/test_lang_builtins_control.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var i interface{} = (*int)(nil)
__check(fmt.Sprint(i == nil), "false") }

// vybe-test: go/lang_builtins_control/untyped_const_shift
// origin: languages/go/tests/go/test_lang_builtins_control.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { const x = 1 << 3
__check(fmt.Sprint(x), "8") }

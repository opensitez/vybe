// vybe-test: go/lang_builtins_control/rune_constant_arithmetic
// origin: languages/go/tests/go/test_lang_builtins_control.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { const c = 'A' + 1
__check(fmt.Sprint(c), "66") }

// vybe-test: go/lang_builtins_control/func_literal_closure
// origin: languages/go/tests/go/test_lang_builtins_control.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { f := func(x int) int { return x + 1 }
__check(fmt.Sprint(f(4)), "5") }

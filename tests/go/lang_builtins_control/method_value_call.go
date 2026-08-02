// vybe-test: go/lang_builtins_control/method_value_call
// origin: languages/go/tests/go/test_lang_builtins_control.rs

package main
import "fmt"
type N int
func (n N) Twice() int { return int(n)*2 }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var n N = 3
f := n.Twice
__check(fmt.Sprint(f()), "6") }

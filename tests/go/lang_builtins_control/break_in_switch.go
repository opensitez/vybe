// vybe-test: go/lang_builtins_control/break_in_switch
// origin: languages/go/tests/go/test_lang_builtins_control.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { switch 1 { case 1: __check(fmt.Sprint("hit"), "hit")
break }
__check(fmt.Sprint("after"), "after") }

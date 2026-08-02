// vybe-test: go/lang_builtins_control/array_value_identity_len
// origin: languages/go/tests/go/test_lang_builtins_control.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var a [3]int
__check(fmt.Sprint(len(a)), "3") }

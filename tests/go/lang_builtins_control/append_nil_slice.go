// vybe-test: go/lang_builtins_control/append_nil_slice
// origin: languages/go/tests/go/test_lang_builtins_control.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var s []int
s = append(s, 1)
__check(fmt.Sprint(len(s)), "1") }

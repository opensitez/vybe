// vybe-test: go/lang_builtins_control/copy_builtin_overlap
// origin: languages/go/tests/go/test_lang_builtins_control.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { s := []int{1,2,3,4}
copy(s[1:], s)
__check(fmt.Sprint(s), "[1 1 2 3]") }

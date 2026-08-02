// vybe-test: go/lang_builtins_control/clear_slice_len_unchanged
// origin: languages/go/tests/go/test_lang_builtins_control.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { s := []int{1,2}
clear(s)
__check(fmt.Sprint(len(s)), "2") }

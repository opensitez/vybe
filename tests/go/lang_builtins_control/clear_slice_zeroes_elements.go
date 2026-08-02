// vybe-test: go/lang_builtins_control/clear_slice_zeroes_elements
// origin: languages/go/tests/go/test_lang_builtins_control.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { s := []int{1,2,3}
clear(s)
__check(fmt.Sprint(s[0]) + " " + fmt.Sprint(s[2]), "0 0") }

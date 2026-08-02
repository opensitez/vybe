// vybe-test: go/lang_builtins_control/slice_from_array_slice_expr
// origin: languages/go/tests/go/test_lang_builtins_control.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { a := [4]int{1,2,3,4}
s := a[1:3]
__check(fmt.Sprint(len(s)) + " " + fmt.Sprint(s[0]), "2 2") }

// vybe-test: go/lang_builtins_control/slice_from_array_full
// origin: languages/go/tests/go/test_lang_builtins_control.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { a := [3]int{1,2,3}
s := a[:]
__check(fmt.Sprint(s[1]), "2") }

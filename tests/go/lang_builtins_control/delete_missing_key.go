// vybe-test: go/lang_builtins_control/delete_missing_key
// origin: languages/go/tests/go/test_lang_builtins_control.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { m := map[string]int{"a":1}
delete(m, "z")
__check(fmt.Sprint(len(m)), "1") }

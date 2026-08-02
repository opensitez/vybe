// vybe-test: go/lang_builtins_control/nil_map_read_zero
// origin: languages/go/tests/go/test_lang_builtins_control.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var m map[string]int
__check(fmt.Sprint(m["x"]), "0") }

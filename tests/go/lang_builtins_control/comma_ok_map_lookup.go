// vybe-test: go/lang_builtins_control/comma_ok_map_lookup
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
_, ok := m["b"]
__check(fmt.Sprint(ok), "false") }

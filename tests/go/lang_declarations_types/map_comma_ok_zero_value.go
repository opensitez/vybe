// vybe-test: go/lang_declarations_types/map_comma_ok_zero_value
// origin: languages/go/tests/go/test_lang_declarations_types.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { m := map[string]int{"a":0}
v, ok := m["a"]
__check(fmt.Sprint(v) + " " + fmt.Sprint(ok), "0 true") }

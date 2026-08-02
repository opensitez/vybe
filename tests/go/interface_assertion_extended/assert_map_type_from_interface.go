// vybe-test: go/interface_assertion_extended/assert_map_type_from_interface
// origin: languages/go/tests/go/test_interface_assertion_extended.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var v interface{} = map[string]int{"a": 1}
m, ok := v.(map[string]int)
__check(fmt.Sprint(m["a"]), "1")
__check(fmt.Sprint(ok), "true") }

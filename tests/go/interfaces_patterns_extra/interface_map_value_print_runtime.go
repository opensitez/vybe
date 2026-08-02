// vybe-test: go/interfaces_patterns_extra/interface_map_value_print_runtime
// origin: languages/go/tests/go/test_interfaces_patterns_extra.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { values := map[string]interface{}{"n": 4}
__check(fmt.Sprint(values["n"]), "4")
}

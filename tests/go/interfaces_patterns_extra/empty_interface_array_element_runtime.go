// vybe-test: go/interfaces_patterns_extra/empty_interface_array_element_runtime
// origin: languages/go/tests/go/test_interfaces_patterns_extra.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { values := [2]interface{}{"a", 2}
__check(fmt.Sprint(values[0]), "a")
__check(fmt.Sprint(values[1]), "2")
}

// vybe-test: go/interfaces_patterns_extra/interface_slice_element_print_runtime
// origin: languages/go/tests/go/test_interfaces_patterns_extra.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { values := []interface{}{1, "go"}
__check(fmt.Sprint(values[1]), "go")
}

// vybe-test: go/interfaces_patterns_extra/interface_assignment_between_variables_runtime
// origin: languages/go/tests/go/test_interfaces_patterns_extra.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var left interface{} = "go"
right := left
__check(fmt.Sprint(right), "go")
}

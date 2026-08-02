// vybe-test: go/interfaces_patterns_extra/interface_reassignment_runtime
// origin: languages/go/tests/go/test_interfaces_patterns_extra.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var value interface{} = 1
__check(fmt.Sprint(value), "1")
value = "two"
__check(fmt.Sprint(value), "two")
}

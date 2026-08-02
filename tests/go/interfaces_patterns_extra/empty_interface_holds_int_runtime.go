// vybe-test: go/interfaces_patterns_extra/empty_interface_holds_int_runtime
// origin: languages/go/tests/go/test_interfaces_patterns_extra.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var value interface{} = 7
__check(fmt.Sprint(value), "7")
}

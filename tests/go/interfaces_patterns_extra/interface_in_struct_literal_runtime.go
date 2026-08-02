// vybe-test: go/interfaces_patterns_extra/interface_in_struct_literal_runtime
// origin: languages/go/tests/go/test_interfaces_patterns_extra.rs

package main
import "fmt"
type holder struct { value interface{} }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { value := holder{value: 11}
__check(fmt.Sprint(value.value), "11")
}

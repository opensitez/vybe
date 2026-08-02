// vybe-test: go/interfaces_patterns_extra/interface_method_returns_string_runtime
// origin: languages/go/tests/go/test_interfaces_patterns_extra.rs

package main
import "fmt"
type namer interface { name() string }
type widget struct{}
func (widget) name() string { return "vybe" }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var value namer = widget{}
__check(fmt.Sprint(value.name()), "vybe")
}

// vybe-test: go/interface_nil_comparable/interface_different_dynamic_types_not_equal
// origin: languages/go/tests/go/test_interface_nil_comparable.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var left interface{} = 1
var right interface{} = "1"
__check(fmt.Sprint(left == right), "false") }

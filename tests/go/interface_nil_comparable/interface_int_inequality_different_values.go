// vybe-test: go/interface_nil_comparable/interface_int_inequality_different_values
// origin: languages/go/tests/go/test_interface_nil_comparable.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var left interface{} = 3
var right interface{} = 4
__check(fmt.Sprint(left == right), "false") }

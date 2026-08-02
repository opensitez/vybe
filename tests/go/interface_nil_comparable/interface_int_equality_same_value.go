// vybe-test: go/interface_nil_comparable/interface_int_equality_same_value
// origin: languages/go/tests/go/test_interface_nil_comparable.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var left interface{} = 7
var right interface{} = 7
__check(fmt.Sprint(left == right), "true") }

// vybe-test: go/interface_nil_comparable/typed_nil_pointers_in_interfaces_equal_same_type
// origin: languages/go/tests/go/test_interface_nil_comparable.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var p *int
var left interface{} = p
var right interface{} = p
__check(fmt.Sprint(left == right), "true") }

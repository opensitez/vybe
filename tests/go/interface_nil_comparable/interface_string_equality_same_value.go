// vybe-test: go/interface_nil_comparable/interface_string_equality_same_value
// origin: languages/go/tests/go/test_interface_nil_comparable.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var left interface{} = "go"
var right interface{} = "go"
__check(fmt.Sprint(left == right), "true") }

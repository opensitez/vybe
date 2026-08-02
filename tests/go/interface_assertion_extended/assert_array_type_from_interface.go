// vybe-test: go/interface_assertion_extended/assert_array_type_from_interface
// origin: languages/go/tests/go/test_interface_assertion_extended.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var v interface{} = [2]int{3, 4}
a, ok := v.([2]int)
__check(fmt.Sprint(a[0] + a[1]), "7")
__check(fmt.Sprint(ok), "true") }

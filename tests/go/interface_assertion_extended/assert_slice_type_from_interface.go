// vybe-test: go/interface_assertion_extended/assert_slice_type_from_interface
// origin: languages/go/tests/go/test_interface_assertion_extended.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var v interface{} = []int{1, 2}
s, ok := v.([]int)
__check(fmt.Sprint(len(s)), "2")
__check(fmt.Sprint(ok), "true") }

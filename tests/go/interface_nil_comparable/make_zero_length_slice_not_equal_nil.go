// vybe-test: go/interface_nil_comparable/make_zero_length_slice_not_equal_nil
// origin: languages/go/tests/go/test_interface_nil_comparable.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { empty := make([]int, 0)
var nilSlice []int
__check(fmt.Sprint(empty == nil), "false")
__check(fmt.Sprint(nilSlice == nil), "true") }

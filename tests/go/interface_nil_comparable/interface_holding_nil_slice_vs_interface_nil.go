// vybe-test: go/interface_nil_comparable/interface_holding_nil_slice_vs_interface_nil
// origin: languages/go/tests/go/test_interface_nil_comparable.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var s []int
var boxed interface{} = s
var empty interface{}
__check(fmt.Sprint(boxed == empty), "false") }

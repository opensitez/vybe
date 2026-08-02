// vybe-test: go/interface_assertion_extended/typed_nil_channel_in_interface_not_nil
// origin: languages/go/tests/go/test_interface_assertion_extended.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var ch chan int
var v interface{} = ch
__check(fmt.Sprint(v == nil), "false") }

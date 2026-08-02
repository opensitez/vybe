// vybe-test: go/interface_nil_comparable/nil_channel_equals_nil_channel
// origin: languages/go/tests/go/test_interface_nil_comparable.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var left chan int
var right chan int
__check(fmt.Sprint(left == nil), "true")
__check(fmt.Sprint(left == right), "true") }

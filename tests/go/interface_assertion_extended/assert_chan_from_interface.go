// vybe-test: go/interface_assertion_extended/assert_chan_from_interface
// origin: languages/go/tests/go/test_interface_assertion_extended.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { ch := make(chan int, 1)
var v interface{} = ch
c, ok := v.(chan int)
c <- 5
__check(fmt.Sprint(<-c), "5")
__check(fmt.Sprint(ok), "true") }

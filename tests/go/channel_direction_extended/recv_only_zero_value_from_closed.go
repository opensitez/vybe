// vybe-test: go/channel_direction_extended/recv_only_zero_value_from_closed
// origin: languages/go/tests/go/test_channel_direction_extended.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { ch := make(chan int)
close(ch)
var r <-chan int = ch
v, ok := <-r
__check(fmt.Sprint(v), "0")
__check(fmt.Sprint(ok), "false") }

// vybe-test: go/channel_direction_extended/bidirectional_close_then_recv_only
// origin: languages/go/tests/go/test_channel_direction_extended.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { ch := make(chan int, 1)
ch <- 21
close(ch)
var r <-chan int = ch
__check(fmt.Sprint(<-r), "21") }

// vybe-test: go/channel_direction_extended/recv_only_after_multiple_sends
// origin: languages/go/tests/go/test_channel_direction_extended.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { ch := make(chan int, 3)
ch <- 1
ch <- 2
ch <- 3
var r <-chan int = ch
__check(fmt.Sprint(<-r + <-r + <-r), "6") }

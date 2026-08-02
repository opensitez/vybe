// vybe-test: go/channel_direction_extended/recv_only_two_receives_fifo
// origin: languages/go/tests/go/test_channel_direction_extended.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { ch := make(chan int, 2)
ch <- 1
ch <- 2
var r <-chan int = ch
__check(fmt.Sprint(<-r), "1")
__check(fmt.Sprint(<-r), "2") }

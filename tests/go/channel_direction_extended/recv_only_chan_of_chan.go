// vybe-test: go/channel_direction_extended/recv_only_chan_of_chan
// origin: languages/go/tests/go/test_channel_direction_extended.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { inner := make(chan int, 1)
inner <- 5
outer := make(chan chan int, 1)
outer <- inner
var r <-chan chan int = outer
ch := <-r
__check(fmt.Sprint(<-ch), "5") }

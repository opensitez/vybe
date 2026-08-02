// vybe-test: go/channel_direction_extended/send_only_chan_of_chan
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
outer := make(chan chan int, 1)
var s chan<- chan int = outer
s <- inner
got := <-outer
got <- 4
__check(fmt.Sprint(<-inner), "4") }

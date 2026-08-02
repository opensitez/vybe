// vybe-test: go/channel_close_range/receive_from_buffered_fifo
// origin: languages/go/tests/go/test_channel_close_range.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { ch := make(chan int, 2)
ch <- 10
ch <- 20
__check(fmt.Sprint(<-ch), "10")
__check(fmt.Sprint(<-ch), "20") }

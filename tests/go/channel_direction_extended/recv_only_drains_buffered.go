// vybe-test: go/channel_direction_extended/recv_only_drains_buffered
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
ch <- 28
ch <- 29
var r <-chan int = ch
__check(fmt.Sprint(<-r), "28")
__check(fmt.Sprint(len(ch)), "1") }

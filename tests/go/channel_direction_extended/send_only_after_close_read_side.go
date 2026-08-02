// vybe-test: go/channel_direction_extended/send_only_after_close_read_side
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
ch <- 1
close(ch)
var r <-chan int = ch
v, ok := <-r
__check(fmt.Sprint(v), "1")
__check(fmt.Sprint(ok), "true") }

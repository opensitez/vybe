// vybe-test: go/channel_direction_extended/receive_only_chan_typed_var
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
ch <- 5
var in <-chan int = ch
__check(fmt.Sprint(<-in), "5") }

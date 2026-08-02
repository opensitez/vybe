// vybe-test: go/channel_direction_extended/assign_bidir_to_recv_only
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
ch <- 9
var r <-chan int = ch
__check(fmt.Sprint(<-r), "9") }

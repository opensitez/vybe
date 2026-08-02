// vybe-test: go/channel_direction_extended/recv_only_string_channel
// origin: languages/go/tests/go/test_channel_direction_extended.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { ch := make(chan string, 1)
ch <- "vybe"
var r <-chan string = ch
__check(fmt.Sprint(<-r), "vybe") }

// vybe-test: go/channel_direction_extended/bidirectional_from_make
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
ch <- "hi"
__check(fmt.Sprint(<-ch), "hi") }

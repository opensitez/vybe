// vybe-test: go/channel_direction_extended/send_only_unbuffered_buffered_fallback
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
var s chan<- int = ch
s <- 9
__check(fmt.Sprint(<-ch), "9") }

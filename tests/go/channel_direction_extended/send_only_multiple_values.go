// vybe-test: go/channel_direction_extended/send_only_multiple_values
// origin: languages/go/tests/go/test_channel_direction_extended.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { ch := make(chan int, 3)
var s chan<- int = ch
s <- 1
s <- 2
s <- 3
__check(fmt.Sprint(<-ch + <-ch + <-ch), "6") }

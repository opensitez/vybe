// vybe-test: go/channel_direction_extended/assign_bidir_to_send_only
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
var s chan<- int = ch
s <- 1
s <- 2
__check(fmt.Sprint(<-ch), "1")
__check(fmt.Sprint(<-ch), "2") }

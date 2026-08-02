// vybe-test: go/channel_direction_extended/send_only_nil_map_not_chan
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
if s != nil { s <- 1 }
__check(fmt.Sprint(<-ch), "1") }

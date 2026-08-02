// vybe-test: go/channel_direction_extended/func_accepts_send_only_and_bidir
// origin: languages/go/tests/go/test_channel_direction_extended.rs

package main
import "fmt"
func fill(ch chan<- int) { ch <- 14 }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { ch := make(chan int, 1)
fill(ch)
__check(fmt.Sprint(<-ch), "14") }

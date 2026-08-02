// vybe-test: go/channel_direction_extended/send_only_passed_to_helper
// origin: languages/go/tests/go/test_channel_direction_extended.rs

package main
import "fmt"
func push(ch chan<- int, v int) { ch <- v }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { ch := make(chan int, 1)
push(ch, 19)
__check(fmt.Sprint(<-ch), "19") }

// vybe-test: go/channel_direction_extended/send_only_in_struct_field
// origin: languages/go/tests/go/test_channel_direction_extended.rs

package main
import "fmt"
type sink struct { ch chan<- int }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { ch := make(chan int, 1)
s := sink{ch: ch}
s.ch <- 5
__check(fmt.Sprint(<-ch), "5") }

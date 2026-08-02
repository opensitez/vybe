// vybe-test: go/channel_direction_extended/recv_only_in_struct_field
// origin: languages/go/tests/go/test_channel_direction_extended.rs

package main
import "fmt"
type source struct { ch <-chan int }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { ch := make(chan int, 1)
ch <- 6
s := source{ch: ch}
__check(fmt.Sprint(<-s.ch), "6") }

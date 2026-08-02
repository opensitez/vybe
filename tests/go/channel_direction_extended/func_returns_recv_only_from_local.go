// vybe-test: go/channel_direction_extended/func_returns_recv_only_from_local
// origin: languages/go/tests/go/test_channel_direction_extended.rs

package main
import "fmt"
func source() <-chan int { ch := make(chan int, 1)
ch <- 13
return ch }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { __check(fmt.Sprint(<-source()), "13") }

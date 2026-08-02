// vybe-test: go/channel_direction_extended/recv_only_param_accepts_bidir
// origin: languages/go/tests/go/test_channel_direction_extended.rs

package main
import "fmt"
func recv(ch <-chan int) int { return <-ch }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { ch := make(chan int, 1)
ch <- 6
__check(fmt.Sprint(recv(ch)), "6") }

// vybe-test: go/channel_select_patterns_extra/channel_receive_after_two_sends_runtime
// origin: languages/go/tests/go/test_channel_select_patterns_extra.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { ch := make(chan int, 2)
ch <- 6
ch <- 7
first := <-ch
second := <-ch
__check(fmt.Sprint(first + second), "13")
}

// vybe-test: go/channel_select_patterns_extra/buffered_channel_len_after_send_runtime
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
ch <- 1
__check(fmt.Sprint(len(ch)), "1")
}

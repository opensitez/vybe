// vybe-test: go/channel_select_patterns_extra/channel_two_sends_len_runtime
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
ch <- 3
ch <- 4
__check(fmt.Sprint(len(ch)), "2")
}

// vybe-test: go/channel_buffered_patterns/buffered_len_after_send
// origin: languages/go/tests/go/test_channel_buffered_patterns.rs

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
__check(fmt.Sprint(len(ch)), "1") }

// vybe-test: go/channel_close_range/buffered_channel_len_cap
// origin: languages/go/tests/go/test_channel_close_range.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { ch := make(chan int, 3)
ch <- 1
ch <- 2
__check(fmt.Sprint(len(ch)), "2")
__check(fmt.Sprint(cap(ch)), "3") }

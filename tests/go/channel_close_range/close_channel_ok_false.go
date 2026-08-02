// vybe-test: go/channel_close_range/close_channel_ok_false
// origin: languages/go/tests/go/test_channel_close_range.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { ch := make(chan int, 1)
ch <- 7
close(ch)
v, ok := <-ch
__check(fmt.Sprint(v), "7")
__check(fmt.Sprint(ok), "true") }

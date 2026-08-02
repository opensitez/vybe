// vybe-test: go/channel_close_range/nil_channel_send_blocks_compile_only
// origin: languages/go/tests/go/test_channel_close_range.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var ch chan int
__check(fmt.Sprint(ch == nil), "true") }

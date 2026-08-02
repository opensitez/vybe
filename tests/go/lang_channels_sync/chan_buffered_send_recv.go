// vybe-test: go/lang_channels_sync/chan_buffered_send_recv
// origin: languages/go/tests/go/test_lang_channels_sync.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { ch := make(chan int, 1)
ch <- 9
__check(fmt.Sprint(<-ch), "9") }

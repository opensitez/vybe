// vybe-test: go/lang_channels_sync/chan_make_unbuffered_zero_cap
// origin: languages/go/tests/go/test_lang_channels_sync.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { ch := make(chan int)
__check(fmt.Sprint(cap(ch)), "0") }

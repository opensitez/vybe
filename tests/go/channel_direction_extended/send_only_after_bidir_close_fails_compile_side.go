// vybe-test: go/channel_direction_extended/send_only_after_bidir_close_fails_compile_side
// origin: languages/go/tests/go/test_channel_direction_extended.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { ch := make(chan int, 1)
close(ch)
var s chan<- int = ch
_, ok := (<-chan int)(ch)
__check(fmt.Sprint(ok), "false") }

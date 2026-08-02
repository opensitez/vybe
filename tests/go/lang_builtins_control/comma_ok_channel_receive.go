// vybe-test: go/lang_builtins_control/comma_ok_channel_receive
// origin: languages/go/tests/go/test_lang_builtins_control.rs

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
_, ok := <-ch
__check(fmt.Sprint(ok), "false") }

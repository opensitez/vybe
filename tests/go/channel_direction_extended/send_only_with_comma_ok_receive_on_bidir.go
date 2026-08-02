// vybe-test: go/channel_direction_extended/send_only_with_comma_ok_receive_on_bidir
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
ch <- 23
v, ok := <-ch
__check(fmt.Sprint(v), "23")
__check(fmt.Sprint(ok), "true") }

// vybe-test: go/channel_direction_extended/func_literal_send_only_param
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
f := func(c chan<- int) { c <- 25 }
f(ch)
__check(fmt.Sprint(<-ch), "25") }

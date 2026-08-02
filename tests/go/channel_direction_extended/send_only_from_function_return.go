// vybe-test: go/channel_direction_extended/send_only_from_function_return
// origin: languages/go/tests/go/test_channel_direction_extended.rs

package main
import "fmt"
func mk() (chan<- int, chan int) { ch := make(chan int, 1)
return ch, ch }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { out, rd := mk()
out <- 11
__check(fmt.Sprint(<-rd), "11") }

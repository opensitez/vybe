// vybe-test: go/channel_direction_extended/recv_only_from_function_param
// origin: languages/go/tests/go/test_channel_direction_extended.rs

package main
import "fmt"
func read(ch <-chan int) { __check(fmt.Sprint(<-ch), "12") }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { ch := make(chan int, 1)
ch <- 12
read(ch) }

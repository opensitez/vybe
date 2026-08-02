// vybe-test: go/channel_direction_extended/recv_only_passed_to_helper
// origin: languages/go/tests/go/test_channel_direction_extended.rs

package main
import "fmt"
func pull(ch <-chan int) int { return <-ch }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { ch := make(chan int, 1)
ch <- 20
__check(fmt.Sprint(pull(ch)), "20") }

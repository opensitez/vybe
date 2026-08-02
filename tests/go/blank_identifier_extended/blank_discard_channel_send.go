// vybe-test: go/blank_identifier_extended/blank_discard_channel_send
// origin: languages/go/tests/go/test_blank_identifier_extended.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { ch := make(chan int, 1)
ch <- 1
_ = <-ch
__check(fmt.Sprint(len(ch)), "0") }

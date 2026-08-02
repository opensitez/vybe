// vybe-test: go/generics_types/generic_chan_wrap_send_recv
// origin: languages/go/tests/go/test_generics_types.rs

package main
import "fmt"
type ChanWrap[T any] struct { Ch chan T }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { w := ChanWrap[int]{Ch: make(chan int, 1)}
w.Ch <- 5
__check(fmt.Sprint(<-w.Ch), "5") }

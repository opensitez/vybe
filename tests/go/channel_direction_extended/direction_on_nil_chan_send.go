// vybe-test: go/channel_direction_extended/direction_on_nil_chan_send
// origin: languages/go/tests/go/test_channel_direction_extended.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { var s chan<- int
__check(fmt.Sprint(s == nil), "true") }

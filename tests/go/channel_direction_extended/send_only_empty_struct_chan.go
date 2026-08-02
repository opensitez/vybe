// vybe-test: go/channel_direction_extended/send_only_empty_struct_chan
// origin: languages/go/tests/go/test_channel_direction_extended.rs

package main
import "fmt"
type sig struct{}
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { ch := make(chan sig, 1)
var s chan<- sig = ch
s <- sig{}
__check(fmt.Sprint("ok"), "ok") }

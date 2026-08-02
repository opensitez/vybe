// vybe-test: go/channel_direction_extended/send_only_bool_channel
// origin: languages/go/tests/go/test_channel_direction_extended.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { ch := make(chan bool, 1)
var s chan<- bool = ch
s <- true
__check(fmt.Sprint(<-ch), "true") }

// vybe-test: go/channel_direction_extended/send_only_pointer_element_type
// origin: languages/go/tests/go/test_channel_direction_extended.rs

package main
import "fmt"
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { ch := make(chan *int, 1)
n := 16
var s chan<- *int = ch
s <- &n
__check(fmt.Sprint(<-*(<-ch)), "16") }

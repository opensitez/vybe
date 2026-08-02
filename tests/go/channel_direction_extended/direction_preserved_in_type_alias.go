// vybe-test: go/channel_direction_extended/direction_preserved_in_type_alias
// origin: languages/go/tests/go/test_channel_direction_extended.rs

package main
import "fmt"
type sendInt chan<- int
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { ch := make(chan int, 1)
var s sendInt = ch
s <- 22
__check(fmt.Sprint(<-ch), "22") }

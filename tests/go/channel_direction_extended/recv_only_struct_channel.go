// vybe-test: go/channel_direction_extended/recv_only_struct_channel
// origin: languages/go/tests/go/test_channel_direction_extended.rs

package main
import "fmt"
type item struct { n int }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { ch := make(chan item, 1)
ch <- item{n: 3}
var r <-chan item = ch
__check(fmt.Sprint((<-r).n), "3") }

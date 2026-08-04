// vybe-test: go/channel_direction_extended/recv_only_with_select_ready
// origin: languages/go/tests/go/test_channel_direction_extended.rs

package main
import "fmt"
var __buf string

// __p appends one line, __pr appends without a newline.
func __p(s string) { __buf = __buf + s + "\n" }

func __pr(s string) { __buf = __buf + s }

// __check ends the program unless the collected output equals want. The final
// Println contributes a trailing newline the expected line vector never
// carried, so both forms are accepted.
func __check(want string) {
	if __buf != want && __buf != want+"\n" {
		fmt.Println("FAIL: want [" + want + "] got [" + __buf + "]")
		panic("assertion failed")
	}
}

func main() { ch := make(chan int, 1)
ch <- 15
var r <-chan int = ch
select { case v := <-r: __p(fmt.Sprint(v))
default: __p(fmt.Sprint(0)) } 
__check("15")
}

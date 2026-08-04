// vybe-test: go/channel_direction_extended/recv_only_param_accepts_bidir
// origin: languages/go/tests/go/test_channel_direction_extended.rs

package main
import "fmt"
func recv(ch <-chan int) int { return <-ch }
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
ch <- 6
__p(fmt.Sprint(recv(ch))) 
__check("6")
}

// vybe-test: go/channel_select_patterns_extra/channel_receive_after_two_sends_runtime
// origin: languages/go/tests/go/test_channel_select_patterns_extra.rs

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

func main() { ch := make(chan int, 2)
ch <- 6
ch <- 7
first := <-ch
second := <-ch
__p(fmt.Sprint(first + second))
__check("13")
}

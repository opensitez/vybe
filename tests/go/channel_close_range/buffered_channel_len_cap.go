// vybe-test: go/channel_close_range/buffered_channel_len_cap
// origin: languages/go/tests/go/test_channel_close_range.rs

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

func main() { ch := make(chan int, 3)
ch <- 1
ch <- 2
__p(fmt.Sprint(len(ch)))
__p(fmt.Sprint(cap(ch))) 
__check("2\n3")
}

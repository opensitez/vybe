// vybe-test: go/channel_direction_extended/send_only_in_struct_field
// origin: languages/go/tests/go/test_channel_direction_extended.rs

package main
import "fmt"
type sink struct { ch chan<- int }
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
s := sink{ch: ch}
s.ch <- 5
__p(fmt.Sprint(<-ch)) 
__check("5")
}

// vybe-test: go/select_patterns_advanced/select_closed_buffered_drains_value_then_zero
// origin: languages/go/tests/go/test_select_patterns_advanced.rs

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
ch <- 42
close(ch)
select { case v := <-ch: __p(fmt.Sprint(v)) }
select { case v, ok := <-ch: __p(fmt.Sprint(v))
__p(fmt.Sprint(ok)) } 
__check("42\n0\nfalse")
}

// vybe-test: go/select_patterns_advanced/select_closed_after_draining_buffered_value
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

func main() { ch := make(chan int, 2)
ch <- 10
ch <- 20
close(ch)
select { case v := <-ch: __p(fmt.Sprint(v)) }
select { case v := <-ch: __p(fmt.Sprint(v)) }
select { case v, ok := <-ch: __p(fmt.Sprint(v))
__p(fmt.Sprint(ok)) } 
__check("10\n20\n0\nfalse")
}

// vybe-test: go/select_patterns_advanced/select_mixed_nil_and_ready_buffered_receive
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

func main() { var blocked chan int
ready := make(chan int, 1)
ready <- 6
select { case <-blocked: __p(fmt.Sprint(0))
case v := <-ready: __p(fmt.Sprint(v))
default: __p(fmt.Sprint("default")) } 
__check("6")
}

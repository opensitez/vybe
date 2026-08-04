// vybe-test: go/select_patterns_advanced/select_string_channel_receive_ready
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

func main() { ch := make(chan string, 1)
ch <- "go"
select { case s := <-ch: __p(fmt.Sprint(s))
default: __p(fmt.Sprint("default")) } 
__check("go")
}

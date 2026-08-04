// vybe-test: go/select_patterns_advanced/select_first_ready_among_three_receive_cases
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

func main() { ch1 := make(chan int, 1)
ch2 := make(chan int)
ch3 := make(chan int)
ch1 <- 3
select { case v := <-ch1: __p(fmt.Sprint(v))
case <-ch2: __p(fmt.Sprint(2))
case <-ch3: __p(fmt.Sprint(1))
default: __p(fmt.Sprint(0)) } 
__check("3")
}

// vybe-test: go/select_patterns_advanced/select_default_with_three_blocked_nil_cases
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

func main() { var a, b, c chan int
select { case <-a: __p(fmt.Sprint(1))
case <-b: __p(fmt.Sprint(2))
case <-c: __p(fmt.Sprint(3))
default: __p(fmt.Sprint("idle")) } 
__check("idle")
}

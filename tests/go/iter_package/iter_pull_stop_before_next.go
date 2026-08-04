// vybe-test: go/iter_package/iter_pull_stop_before_next
// origin: languages/go/tests/go/test_iter_package.rs

package main
import "fmt"
import "iter"
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

func main() { ran := 0
seq := func(yield func(int) bool) { ran++
yield(1)
yield(2) }
next, stop := iter.Pull(seq)
stop()
__p(fmt.Sprint(ran)) 
__check("0")
}

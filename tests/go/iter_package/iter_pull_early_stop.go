// vybe-test: go/iter_package/iter_pull_early_stop
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

func main() { seq := func(yield func(int) bool) { if !yield(10) { return }
yield(20) }
next, stop := iter.Pull(seq)
defer stop()
v, ok := next()
stop()
__p(fmt.Sprint(v))
__p(fmt.Sprint(ok)) 
__check("10\ntrue")
}

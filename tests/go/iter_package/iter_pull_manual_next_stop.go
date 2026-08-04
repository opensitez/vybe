// vybe-test: go/iter_package/iter_pull_manual_next_stop
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

func main() { seq := func(yield func(int) bool) { yield(1)
yield(2)
yield(3) }
next, stop := iter.Pull(seq)
defer stop()
v1, ok1 := next()
v2, ok2 := next()
_, ok3 := next()
_, ok4 := next()
__p(fmt.Sprint(v1))
__p(fmt.Sprint(v2))
__p(fmt.Sprint(ok1 && ok2 && ok3 && !ok4)) 
__check("1\n2\ntrue")
}

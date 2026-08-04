// vybe-test: go/iter_package/iter_pull2_exhausted
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

func main() { seq := func(yield func(int, int) bool) { yield(0, 0) }
next, stop := iter.Pull2(seq)
defer stop()
_, _, ok1 := next()
_, _, ok2 := next()
__p(fmt.Sprint(ok1))
__p(fmt.Sprint(ok2)) 
__check("true\nfalse")
}

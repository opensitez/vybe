// vybe-test: go/iter_package/iter_pull_three_values
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

func main() { seq := func(yield func(int) bool) { yield(5)
yield(6)
yield(7) }
next, stop := iter.Pull(seq)
defer stop()
a, _ := next()
b, _ := next()
c, _ := next()
__p(fmt.Sprint(a + b + c)) 
__check("18")
}

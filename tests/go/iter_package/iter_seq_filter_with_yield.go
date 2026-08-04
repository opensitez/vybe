// vybe-test: go/iter_package/iter_seq_filter_with_yield
// origin: languages/go/tests/go/test_iter_package.rs

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

func main() { seq := func(yield func(int) bool) { for i := 1; i <= 5; i++ { if i%2 == 0 { if !yield(i) { return } } } }
evens := 0
for range seq { evens++ }
__p(fmt.Sprint(evens)) 
__check("2")
}

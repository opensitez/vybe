// vybe-test: go/declarations_patterns/tuple_assignment_from_multi_return_runtime
// origin: languages/go/tests/go/test_declarations_patterns.rs

package main
import "fmt"
func dims() (int, int) { return 4, 6 }
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

func main() { w, h := dims()
__p(fmt.Sprint(w * h))
__check("24")
}

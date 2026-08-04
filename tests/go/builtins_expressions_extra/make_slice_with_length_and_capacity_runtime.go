// vybe-test: go/builtins_expressions_extra/make_slice_with_length_and_capacity_runtime
// origin: languages/go/tests/go/test_builtins_expressions_extra.rs

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

func main() { values := make([]int, 3, 6)
__p(fmt.Sprint(len(values)))
__p(fmt.Sprint(cap(values)))
__check("3\n6")
}

// vybe-test: go/range_iteration_extra/range_over_slice_modify_by_index_runtime
// origin: languages/go/tests/go/test_range_iteration_extra.rs

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

func main() { values := []int{1, 2, 3}
for index := range values { values[index]++ }
__p(fmt.Sprint(values[0]))
__p(fmt.Sprint(values[2]))
__check("2\n4")
}

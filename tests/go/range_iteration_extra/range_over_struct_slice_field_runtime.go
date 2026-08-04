// vybe-test: go/range_iteration_extra/range_over_struct_slice_field_runtime
// origin: languages/go/tests/go/test_range_iteration_extra.rs

package main
import "fmt"
type holder struct { values []int }
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

func main() { value := holder{values: []int{1, 2, 3}}
total := 0
for _, item := range value.values { total += item }
__p(fmt.Sprint(total))
__check("6")
}

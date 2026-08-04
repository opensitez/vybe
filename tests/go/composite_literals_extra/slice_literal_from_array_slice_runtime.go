// vybe-test: go/composite_literals_extra/slice_literal_from_array_slice_runtime
// origin: languages/go/tests/go/test_composite_literals_extra.rs

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

func main() { values := [4]int{1, 2, 3, 4}
part := values[1:3]
__p(fmt.Sprint(part[0]))
__p(fmt.Sprint(part[1]))
__check("2\n3")
}

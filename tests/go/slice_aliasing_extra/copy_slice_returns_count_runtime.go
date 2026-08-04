// vybe-test: go/slice_aliasing_extra/copy_slice_returns_count_runtime
// origin: languages/go/tests/go/test_slice_aliasing_extra.rs

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

func main() { dst := make([]int, 2)
src := []int{3, 4, 5}
__p(fmt.Sprint(copy(dst, src)))
__check("2")
}

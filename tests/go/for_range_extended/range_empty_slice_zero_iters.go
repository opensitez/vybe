// vybe-test: go/for_range_extended/range_empty_slice_zero_iters
// origin: languages/go/tests/go/test_for_range_extended.rs

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

func main() { count := 0
for _, v := range []int{} { count += v }
__p(fmt.Sprint(count)) 
__check("0")
}

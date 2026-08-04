// vybe-test: go/for_range_extended/range_slice_reassign_loop_var
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

func main() { sum := 0
for _, v := range []int{1, 2, 3} { v = v * 10
sum += v }
__p(fmt.Sprint(sum)) 
__check("60")
}

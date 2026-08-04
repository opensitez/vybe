// vybe-test: go/variadic_advanced/spread_nil_slice_len_zero
// origin: languages/go/tests/go/test_variadic_advanced.rs

package main
import "fmt"
func size(items ...int) int { return len(items) }
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

func main() { var s []int
__p(fmt.Sprint(size(s...))) 
__check("0")
}

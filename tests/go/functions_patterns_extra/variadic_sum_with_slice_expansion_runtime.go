// vybe-test: go/functions_patterns_extra/variadic_sum_with_slice_expansion_runtime
// origin: languages/go/tests/go/test_functions_patterns_extra.rs

package main
import "fmt"
func sum(values ...int) int { total := 0
for _, v := range values { total += v }
return total }
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

func main() { nums := []int{4, 5, 6}
__p(fmt.Sprint(sum(nums...)))
__check("15")
}

// vybe-test: go/variadic/variadic_max
// origin: languages/go/tests/go/test_variadic.rs

package main
import "fmt"
func max(nums ...int) int { m := nums[0]
for _, n := range nums { if n > m { m = n } }
return m } var __buf string

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

func main() { __p(fmt.Sprint(max(3, 1, 4, 1, 5, 9, 2)))
__check("9")
}

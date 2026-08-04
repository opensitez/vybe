// vybe-test: go/variadic_spread/variadic_int_minimum
// origin: languages/go/tests/go/test_variadic_spread.rs

package main
import "fmt"
func minimum(nums ...int) int { m := nums[0]
for _, n := range nums { if n < m { m = n } }
return m }
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

func main() { __p(fmt.Sprint(minimum(5, 1, 8, 2)))
__check("1")
}

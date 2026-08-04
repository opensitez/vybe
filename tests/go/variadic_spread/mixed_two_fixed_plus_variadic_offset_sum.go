// vybe-test: go/variadic_spread/mixed_two_fixed_plus_variadic_offset_sum
// origin: languages/go/tests/go/test_variadic_spread.rs

package main
import "fmt"
func tally(base int, step int, nums ...int) int { total := base
for _, n := range nums { total += n + step }
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

func main() { __p(fmt.Sprint(tally(100, 1, 2, 3)))
__check("108")
}

// vybe-test: go/variadic_spread/forward_variadic_injects_count_before_delegate
// origin: languages/go/tests/go/test_variadic_spread.rs

package main
import "fmt"
func emit(nums ...int) { for _, n := range nums { __p(fmt.Sprint(n)) } }
func relay(nums ...int) { __p(fmt.Sprint(len(nums)))
emit(nums...) }
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

func main() { relay(7, 8)
__check("2\n7\n8")
}

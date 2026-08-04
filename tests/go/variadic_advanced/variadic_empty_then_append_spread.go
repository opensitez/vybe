// vybe-test: go/variadic_advanced/variadic_empty_then_append_spread
// origin: languages/go/tests/go/test_variadic_advanced.rs

package main
import "fmt"
func lenAfter(base []int, more ...int) int { combined := append(base, more...)
return len(combined) }
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

func main() { __p(fmt.Sprint(lenAfter([]int{1}, 2, 3))) 
__check("3")
}

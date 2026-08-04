// vybe-test: go/defer_lifo_extended/defer_named_return_two_values_both_set
// origin: languages/go/tests/go/test_defer_lifo_extended.rs

package main
import "fmt"
func work() (a int, b int) { defer func() { a = 1
b = 2 }()
return 9, 8 }
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

func main() { x, y := work()
__p(fmt.Sprint(x))
__p(fmt.Sprint(y)) 
__check("1\n2")
}

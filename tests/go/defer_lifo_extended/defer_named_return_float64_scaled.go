// vybe-test: go/defer_lifo_extended/defer_named_return_float64_scaled
// origin: languages/go/tests/go/test_defer_lifo_extended.rs

package main
import "fmt"
func work() (f float64) { defer func() { f = f * 2 }()
return 3.5 }
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

func main() { __p(fmt.Sprint(f == 7.0)) 
__check("true")
}

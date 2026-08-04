// vybe-test: go/range_over_int/range_int_function_return_bound
// origin: languages/go/tests/go/test_range_over_int.rs

package main
import "fmt"
func bound() int { return 4 }
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

func main() { total := 0
for i := range bound() { total += i }
__p(fmt.Sprint(total)) 
__check("6")
}

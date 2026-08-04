// vybe-test: go/function_types_advanced/method_apply_twice_with_predicate_param
// origin: languages/go/tests/go/test_function_types_advanced.rs

package main
import "fmt"
type tally struct { count int }
func (t *tally) whilePositive(ok func(int) bool) { for ok(t.count) { t.count-- } }
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

func main() { value := tally{count: 3}
value.whilePositive(func(v int) bool { return v > 0 })
__p(fmt.Sprint(value.count)) 
__check("0")
}

// vybe-test: go/method_sets_pointer_value/value_with_only_pointer_methods_needs_address_runtime
// origin: languages/go/tests/go/test_method_sets_pointer_value.rs

package main
import "fmt"
type latch struct { on bool }
func (l *latch) flip() { l.on = !l.on }
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

func main() { l := latch{on: false}
l.flip()
__p(fmt.Sprint(l.on)) 
__check("true")
}

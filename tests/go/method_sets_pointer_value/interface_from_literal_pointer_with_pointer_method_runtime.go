// vybe-test: go/method_sets_pointer_value/interface_from_literal_pointer_with_pointer_method_runtime
// origin: languages/go/tests/go/test_method_sets_pointer_value.rs

package main
import "fmt"
type resetter interface { reset() }
type timer struct { ticks int }
func (t *timer) reset() { t.ticks = 0 }
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

func main() { var r resetter = &timer{ticks: 5}
r.reset()
__p(fmt.Sprint(r.(*timer).ticks)) 
__check("0")
}

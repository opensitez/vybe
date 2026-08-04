// vybe-test: go/method_sets_pointer_value/embedded_value_type_pointer_method_on_outer_value_runtime
// origin: languages/go/tests/go/test_method_sets_pointer_value.rs

package main
import "fmt"
type engine struct { rpm int }
func (e *engine) rev() { e.rpm++ }
type car struct { engine }
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

func main() { c := car{engine: engine{rpm: 1000}}
c.rev()
__p(fmt.Sprint(c.rpm)) 
__check("1001")
}

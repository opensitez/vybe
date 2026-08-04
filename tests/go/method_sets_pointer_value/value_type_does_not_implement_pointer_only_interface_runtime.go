// vybe-test: go/method_sets_pointer_value/value_type_does_not_implement_pointer_only_interface_runtime
// origin: languages/go/tests/go/test_method_sets_pointer_value.rs

package main
import "fmt"
type mutator interface { set(int) }
type gauge struct { n int }
func (g *gauge) set(v int) { g.n = v }
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

func main() { g := gauge{}
var m mutator = &g
m.set(4)
__p(fmt.Sprint(g.n)) 
__check("4")
}

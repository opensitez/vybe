// vybe-test: go/interface_embedding_methods/overlapping_method_unified_impl_runtime
// origin: languages/go/tests/go/test_interface_embedding_methods.rs

package main
import "fmt"
type resetterA interface { reset() int }
type resetterB interface { reset() int }
type dualReset interface { resetterA
resetterB }
type engine struct { ticks int }
func (e *engine) reset() int { e.ticks = 0
return e.ticks }
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

func main() { value := &engine{ticks: 5}
var d dualReset = value
__p(fmt.Sprint(d.reset())) 
__check("0")
}

// vybe-test: go/interface_embedding_methods/composite_interface_in_slice_runtime
// origin: languages/go/tests/go/test_interface_embedding_methods.rs

package main
import "fmt"
type mover interface { move() int }
type jumper interface { mover }
type hopper struct { steps int }
func (h hopper) move() int { return h.steps }
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

func main() { values := []jumper{hopper{steps: 3}}
__p(fmt.Sprint(values[0].move())) 
__check("3")
}

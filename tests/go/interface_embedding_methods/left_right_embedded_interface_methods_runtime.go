// vybe-test: go/interface_embedding_methods/left_right_embedded_interface_methods_runtime
// origin: languages/go/tests/go/test_interface_embedding_methods.rs

package main
import "fmt"
type left interface { side() string }
type right interface { edge() string }
type pair interface { left
right }
type both struct{}
func (both) side() string { return "L" }
func (both) edge() string { return "R" }
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

func main() { var p pair = both{}
__p(fmt.Sprint(p.side()))
__p(fmt.Sprint(p.edge())) 
__check("L\nR")
}

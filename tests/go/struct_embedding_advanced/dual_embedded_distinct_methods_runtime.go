// vybe-test: go/struct_embedding_advanced/dual_embedded_distinct_methods_runtime
// origin: languages/go/tests/go/test_struct_embedding_advanced.rs

package main
import "fmt"
type left struct{}
func (left) side() string { return "L" }
type right struct{}
func (right) edge() string { return "R" }
type pair struct { left
right }
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

func main() { value := pair{}
__p(fmt.Sprint(value.side()))
__p(fmt.Sprint(value.edge()))
__check("L\nR")
}

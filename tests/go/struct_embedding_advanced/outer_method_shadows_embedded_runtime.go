// vybe-test: go/struct_embedding_advanced/outer_method_shadows_embedded_runtime
// origin: languages/go/tests/go/test_struct_embedding_advanced.rs

package main
import "fmt"
type inner struct{}
func (inner) label() string { return "inner" }
type outer struct { inner }
func (outer) label() string { return "outer" }
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

func main() { value := outer{}
__p(fmt.Sprint(value.label()))
__check("outer")
}

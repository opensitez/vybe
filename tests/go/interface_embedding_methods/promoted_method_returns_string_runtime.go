// vybe-test: go/interface_embedding_methods/promoted_method_returns_string_runtime
// origin: languages/go/tests/go/test_interface_embedding_methods.rs

package main
import "fmt"
type namer interface { name() string }
type labeled interface { namer }
type widget struct{}
func (widget) name() string { return "vybe" }
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

func main() { var value labeled = widget{}
__p(fmt.Sprint(value.name())) 
__check("vybe")
}

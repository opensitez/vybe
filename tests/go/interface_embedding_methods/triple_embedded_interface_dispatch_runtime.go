// vybe-test: go/interface_embedding_methods/triple_embedded_interface_dispatch_runtime
// origin: languages/go/tests/go/test_interface_embedding_methods.rs

package main
import "fmt"
type leaf interface { tag() string }
type branch interface { leaf }
type trunk interface { branch }
type node struct{}
func (node) tag() string { return "deep" }
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

func main() { var t trunk = node{}
__p(fmt.Sprint(t.tag())) 
__check("deep")
}

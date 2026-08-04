// vybe-test: go/interface_embedding_methods/promoted_interface_method_value_runtime
// origin: languages/go/tests/go/test_interface_embedding_methods.rs

package main
import "fmt"
type greeter interface { greet() string }
type social interface { greeter }
type hi struct{}
func (hi) greet() string { return "hi" }
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

func main() { var s social = hi{}
fn := s.greet
__p(fmt.Sprint(fn())) 
__check("hi")
}

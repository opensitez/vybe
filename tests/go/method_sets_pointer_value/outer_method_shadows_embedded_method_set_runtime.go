// vybe-test: go/method_sets_pointer_value/outer_method_shadows_embedded_method_set_runtime
// origin: languages/go/tests/go/test_method_sets_pointer_value.rs

package main
import "fmt"
type base struct{}
func (base) tag() string { return "base" }
type derived struct { base }
func (derived) tag() string { return "derived" }
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

func main() { d := derived{}
__p(fmt.Sprint(d.tag())) 
__check("derived")
}

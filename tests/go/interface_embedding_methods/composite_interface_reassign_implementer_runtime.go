// vybe-test: go/interface_embedding_methods/composite_interface_reassign_implementer_runtime
// origin: languages/go/tests/go/test_interface_embedding_methods.rs

package main
import "fmt"
type sized interface { size() int }
type measurable interface { sized }
type small struct{}
func (small) size() int { return 1 }
type large struct{}
func (large) size() int { return 9 }
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

func main() { var m measurable = small{}
__p(fmt.Sprint(m.size()))
m = large{}
__p(fmt.Sprint(m.size())) 
__check("1\n9")
}

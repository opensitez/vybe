// vybe-test: go/interface_assertion_extended/typed_nil_pointer_in_named_interface_not_nil
// origin: languages/go/tests/go/test_interface_assertion_extended.rs

package main
import "fmt"
type holder interface { size() int }
type box struct { n int }
func (b *box) size() int { return b.n }
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

func main() { var p *box
var h holder = p
__p(fmt.Sprint(h == nil)) 
__check("false")
}

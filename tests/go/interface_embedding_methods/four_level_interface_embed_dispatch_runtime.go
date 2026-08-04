// vybe-test: go/interface_embedding_methods/four_level_interface_embed_dispatch_runtime
// origin: languages/go/tests/go/test_interface_embedding_methods.rs

package main
import "fmt"
type d interface { n() int }
type c interface { d }
type b interface { c }
type a interface { b }
type leaf struct { value int }
func (l leaf) n() int { return l.value }
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

func main() { var top a = leaf{value: 13}
__p(fmt.Sprint(top.n())) 
__check("13")
}

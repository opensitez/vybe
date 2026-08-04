// vybe-test: go/interface_embedding_methods/composite_interface_struct_impl_runtime
// origin: languages/go/tests/go/test_interface_embedding_methods.rs

package main
import "fmt"
type opener interface { open() bool }
type closer interface { close() }
type resource interface { opener
closer }
type file struct { ok bool }
func (f file) open() bool { return f.ok }
func (f file) close() {}
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

func main() { var r resource = file{ok: true}
__p(fmt.Sprint(r.open())) 
__check("true")
}

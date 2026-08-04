// vybe-test: go/method_sets_pointer_value/pointer_type_satisfies_interface_with_pointer_method_runtime
// origin: languages/go/tests/go/test_method_sets_pointer_value.rs

package main
import "fmt"
type writer interface { write(int) }
type pad struct { n int }
func (p *pad) write(v int) { p.n = v }
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

func main() { var w writer = &pad{}
w.write(9)
__p(fmt.Sprint(w.(*pad).n)) 
__check("9")
}

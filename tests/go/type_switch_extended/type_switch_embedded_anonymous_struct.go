// vybe-test: go/type_switch_extended/type_switch_embedded_anonymous_struct
// origin: languages/go/tests/go/test_type_switch_extended.rs

package main
import "fmt"
type base struct { id int }
type child struct { base
name string }
func tag(v interface{}) { switch t := v.(type) { case child: __p(fmt.Sprint(t.id))
__p(fmt.Sprint(t.name))
default: __p(fmt.Sprint("x")) } }
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

func main() { tag(child{base: base{id: 3}, name: "c"}) 
__check("3\nc")
}

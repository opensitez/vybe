// vybe-test: go/type_switch_extended/type_switch_nested_struct_field
// origin: languages/go/tests/go/test_type_switch_extended.rs

package main
import "fmt"
type inner struct { x int }
type outer struct { in inner }
func tag(v interface{}) { switch t := v.(type) { case outer: __p(fmt.Sprint(t.in.x))
default: __p(fmt.Sprint(0)) } }
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

func main() { tag(outer{in: inner{x: 8}}) 
__check("8")
}

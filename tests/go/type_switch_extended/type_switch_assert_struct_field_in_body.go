// vybe-test: go/type_switch_extended/type_switch_assert_struct_field_in_body
// origin: languages/go/tests/go/test_type_switch_extended.rs

package main
import "fmt"
type pair struct { a int
b int }
func work(v interface{}) { switch x := v.(type) { case pair: __p(fmt.Sprint(x.a + x.b))
case *pair: __p(fmt.Sprint(x.b))
default: __p(fmt.Sprint(-1)) } }
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

func main() { work(pair{a: 2, b: 5}) 
__check("7")
}

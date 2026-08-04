// vybe-test: go/type_switch_extended/type_switch_interface_subcase_via_type_assert
// origin: languages/go/tests/go/test_type_switch_extended.rs

package main
import "fmt"
type fmtStringer interface { String() string }
type myInt int
func (m myInt) String() string { return "m" }
func tag(v interface{}) { switch v.(type) { case fmtStringer: __p(fmt.Sprint(v.(fmtStringer).String()))
default: __p(fmt.Sprint("no")) } }
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

func main() { tag(myInt(1)) 
__check("m")
}

// vybe-test: go/pointer_receivers_advanced/new_struct_zero_value_matches_literal_runtime
// origin: languages/go/tests/go/test_pointer_receivers_advanced.rs

package main
import "fmt"
type widget struct { size int
label string }
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

func main() { fromNew := new(widget)
fromLit := &widget{}
__p(fmt.Sprint(fromNew.size == fromLit.size))
__p(fmt.Sprint(fromNew.label == fromLit.label))
__check("true\ntrue")
}

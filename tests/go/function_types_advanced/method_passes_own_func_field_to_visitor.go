// vybe-test: go/function_types_advanced/method_passes_own_func_field_to_visitor
// origin: languages/go/tests/go/test_function_types_advanced.rs

package main
import "fmt"
type node struct { label string
format func(string) string }
func (n node) show(visitor func(string)) { visitor(n.format(n.label)) }
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

func main() { value := node{label: "go", format: func(s string) string { return s + "!" }}
value.show(func(s string) { __p(fmt.Sprint(s)) }) 
__check("go!")
}

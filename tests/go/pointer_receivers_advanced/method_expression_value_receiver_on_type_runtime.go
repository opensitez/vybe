// vybe-test: go/pointer_receivers_advanced/method_expression_value_receiver_on_type_runtime
// origin: languages/go/tests/go/test_pointer_receivers_advanced.rs

package main
import "fmt"
type tag struct { name string }
func (t tag) label() string { return t.name }
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

func main() { fn := tag.label
__p(fmt.Sprint(fn(tag{name: "go"})))
__check("go")
}

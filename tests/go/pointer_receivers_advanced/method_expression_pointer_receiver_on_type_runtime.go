// vybe-test: go/pointer_receivers_advanced/method_expression_pointer_receiver_on_type_runtime
// origin: languages/go/tests/go/test_pointer_receivers_advanced.rs

package main
import "fmt"
type tag struct { name string }
func (t *tag) rename(v string) { t.name = v }
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

func main() { value := &tag{name: "old"}
fn := (*tag).rename
fn(value, "new")
__p(fmt.Sprint(value.name))
__check("new")
}

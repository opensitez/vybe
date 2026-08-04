// vybe-test: go/pointer_receivers_advanced/value_receiver_slice_append_not_visible_runtime
// origin: languages/go/tests/go/test_pointer_receivers_advanced.rs

package main
import "fmt"
type bag struct { items []int }
func (b bag) push(v int) { b.items = append(b.items, v) }
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

func main() { value := bag{items: []int{1}}
value.push(2)
__p(fmt.Sprint(len(value.items)))
__check("1")
}

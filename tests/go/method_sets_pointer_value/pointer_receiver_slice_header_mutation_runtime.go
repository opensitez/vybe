// vybe-test: go/method_sets_pointer_value/pointer_receiver_slice_header_mutation_runtime
// origin: languages/go/tests/go/test_method_sets_pointer_value.rs

package main
import "fmt"
type bag struct { items []int }
func (b *bag) appendItem(v int) { b.items = append(b.items, v) }
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

func main() { b := bag{items: []int{1}}
b.appendItem(2)
__p(fmt.Sprint(len(b.items)))
__p(fmt.Sprint(b.items[1])) 
__check("2\n2")
}

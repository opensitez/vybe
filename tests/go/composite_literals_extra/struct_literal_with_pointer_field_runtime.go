// vybe-test: go/composite_literals_extra/struct_literal_with_pointer_field_runtime
// origin: languages/go/tests/go/test_composite_literals_extra.rs

package main
import "fmt"
type node struct { value int }
type wrapper struct { item *node }
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

func main() { n := node{value: 12}
w := wrapper{item: &n}
__p(fmt.Sprint(w.item.value))
__check("12")
}

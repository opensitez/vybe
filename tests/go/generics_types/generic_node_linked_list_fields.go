// vybe-test: go/generics_types/generic_node_linked_list_fields
// origin: languages/go/tests/go/test_generics_types.rs

package main
import "fmt"
type Node[T any] struct { Val T
Next *Node[T] }
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

func main() { head := &Node[int]{Val: 1, Next: &Node[int]{Val: 2}}
__p(fmt.Sprint(head.Val))
__p(fmt.Sprint(head.Next.Val)) 
__check("1\n2")
}

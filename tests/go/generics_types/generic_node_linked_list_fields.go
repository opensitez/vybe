// vybe-test: go/generics_types/generic_node_linked_list_fields
// origin: languages/go/tests/go/test_generics_types.rs

package main
import "fmt"
type Node[T any] struct { Val T
Next *Node[T] }
func __check(got string, want string) {
	if got != want {
		fmt.Println("FAIL: want [" + want + "] got [" + got + "]")
		panic("assertion failed")
	}
}

func main() { head := &Node[int]{Val: 1, Next: &Node[int]{Val: 2}}
__check(fmt.Sprint(head.Val), "1")
__check(fmt.Sprint(head.Next.Val), "2") }

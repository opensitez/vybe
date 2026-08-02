// vybe-test: go/generics_constraints_extended/generic_method_pointer_on_generic_struct
// origin: languages/go/tests/go/test_generics_constraints_extended.rs
// vybe-test-mode: compile

package main
type Node[T any] struct { Val T
Next *Node[T] }
func (n *Node[T]) Link(next *Node[T]) { n.Next = next }
func main() { a, b := &Node[int]{}, &Node[int]{}
a.Link(b) }

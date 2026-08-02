// vybe-test: go/generics_types/generic_method_chained_on_pointer
// origin: languages/go/tests/go/test_generics_types.rs
// vybe-test-mode: compile

package main
type Node[T any] struct { Val T
Next *Node[T] }
func (n *Node[T]) Link(next *Node[T]) *Node[T] { n.Next = next
return n }
func main() { _ = (&Node[int]{Val: 1}).Link(&Node[int]{Val: 2}) }

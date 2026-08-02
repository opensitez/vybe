// vybe-test: go/unsafe_size_align_extended/no_arithmetic_pointer_to_struct
// origin: languages/go/tests/go/test_unsafe_size_align_extended.rs
// vybe-test-mode: compile

package main
import "unsafe"
type Node struct { next *Node }
func main() { var n Node
_ = unsafe.Pointer(&n) }

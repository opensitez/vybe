// vybe-test: go/lang_declarations_types/recursive_struct_via_pointer
// origin: languages/go/tests/go/test_lang_declarations_types.rs
// vybe-test-mode: compile

package main
type Node struct { Next *Node }
func main() { _ = Node{} }

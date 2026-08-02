// vybe-test: go/struct_tags_reflect/struct_tag_on_pointer_field_compile
// origin: languages/go/tests/go/test_struct_tags_reflect.rs
// vybe-test-mode: compile

package main
type Node struct { Next *Node `json:"next,omitempty"` }
func main() { _ = Node{} }

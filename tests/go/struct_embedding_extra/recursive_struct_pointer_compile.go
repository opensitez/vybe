// vybe-test: go/struct_embedding_extra/recursive_struct_pointer_compile
// origin: languages/go/tests/go/test_struct_embedding_extra.rs
// vybe-test-mode: compile

package main
type node struct { next *node }
func main() { var n node
_ = n }

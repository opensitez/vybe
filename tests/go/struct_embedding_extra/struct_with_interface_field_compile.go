// vybe-test: go/struct_embedding_extra/struct_with_interface_field_compile
// origin: languages/go/tests/go/test_struct_embedding_extra.rs
// vybe-test-mode: compile

package main
type holder struct { value interface{} }
func main() { _ = holder{} }

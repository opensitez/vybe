// vybe-test: go/struct_embedding_extra/struct_with_function_field_compile
// origin: languages/go/tests/go/test_struct_embedding_extra.rs
// vybe-test-mode: compile

package main
type holder struct { fn func(int) int }
func main() { _ = holder{} }

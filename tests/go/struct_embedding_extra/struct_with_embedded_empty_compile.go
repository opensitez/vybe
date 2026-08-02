// vybe-test: go/struct_embedding_extra/struct_with_embedded_empty_compile
// origin: languages/go/tests/go/test_struct_embedding_extra.rs
// vybe-test-mode: compile

package main
type marker struct{}
type holder struct { marker }
func main() { _ = holder{} }

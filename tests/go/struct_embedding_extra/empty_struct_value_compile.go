// vybe-test: go/struct_embedding_extra/empty_struct_value_compile
// origin: languages/go/tests/go/test_struct_embedding_extra.rs
// vybe-test-mode: compile

package main
type marker struct{}
func main() { _ = marker{} }

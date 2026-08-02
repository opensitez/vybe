// vybe-test: go/struct_embedding_extra/struct_with_map_field_compile
// origin: languages/go/tests/go/test_struct_embedding_extra.rs
// vybe-test-mode: compile

package main
type bag struct { values map[string]int }
func main() { _ = bag{} }

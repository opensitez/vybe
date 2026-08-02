// vybe-test: go/struct_embedding_extra/struct_nested_literal_compile
// origin: languages/go/tests/go/test_struct_embedding_extra.rs
// vybe-test-mode: compile

package main
type inner struct { count int }
type outer struct { value inner }
func main() { _ = outer{value: inner{count: 3}} }

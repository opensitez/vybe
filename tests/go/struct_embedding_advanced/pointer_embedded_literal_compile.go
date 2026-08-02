// vybe-test: go/struct_embedding_advanced/pointer_embedded_literal_compile
// origin: languages/go/tests/go/test_struct_embedding_advanced.rs
// vybe-test-mode: compile

package main
type inner struct { count int }
type outer struct { *inner }
func main() { _ = outer{inner: &inner{count: 1}} }

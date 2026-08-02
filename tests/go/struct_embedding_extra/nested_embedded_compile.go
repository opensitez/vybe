// vybe-test: go/struct_embedding_extra/nested_embedded_compile
// origin: languages/go/tests/go/test_struct_embedding_extra.rs
// vybe-test-mode: compile

package main
type inner struct { count int }
type middle struct { inner }
type outer struct { middle }
func main() { var value outer
_ = value.count }

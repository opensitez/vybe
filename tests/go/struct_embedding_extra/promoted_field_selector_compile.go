// vybe-test: go/struct_embedding_extra/promoted_field_selector_compile
// origin: languages/go/tests/go/test_struct_embedding_extra.rs
// vybe-test-mode: compile

package main
type inner struct { count int }
type outer struct { inner }
func main() { var value outer
_ = value.count }

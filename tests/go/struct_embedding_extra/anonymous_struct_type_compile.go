// vybe-test: go/struct_embedding_extra/anonymous_struct_type_compile
// origin: languages/go/tests/go/test_struct_embedding_extra.rs
// vybe-test-mode: compile

package main
func main() { var value struct { count int
label string }
_ = value }

// vybe-test: go/lang_interfaces_embedding/struct_anonymous_field_name
// origin: languages/go/tests/go/test_lang_interfaces_embedding.rs
// vybe-test-mode: compile

package main
type T struct { int }
func main() { _ = T{int: 1} }

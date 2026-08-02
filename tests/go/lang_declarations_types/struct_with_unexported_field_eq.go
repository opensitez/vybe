// vybe-test: go/lang_declarations_types/struct_with_unexported_field_eq
// origin: languages/go/tests/go/test_lang_declarations_types.rs
// vybe-test-mode: compile

package main
type t struct { x int }
func main() { _ = t{} == t{} }

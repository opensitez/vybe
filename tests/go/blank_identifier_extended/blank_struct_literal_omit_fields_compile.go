// vybe-test: go/blank_identifier_extended/blank_struct_literal_omit_fields_compile
// origin: languages/go/tests/go/test_blank_identifier_extended.rs
// vybe-test-mode: compile

package main
type node struct { id int
name string }
func main() { n := node{id: 1}
_ = n.id }

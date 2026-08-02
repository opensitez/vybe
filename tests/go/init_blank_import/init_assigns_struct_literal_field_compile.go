// vybe-test: go/init_blank_import/init_assigns_struct_literal_field_compile
// origin: languages/go/tests/go/test_init_blank_import.rs
// vybe-test-mode: compile

package main
type node struct { value int }
var root node
func init() { root = node{value: 9} }
func main() { _ = root.value }

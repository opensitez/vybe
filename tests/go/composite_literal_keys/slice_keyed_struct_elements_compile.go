// vybe-test: go/composite_literal_keys/slice_keyed_struct_elements_compile
// origin: languages/go/tests/go/test_composite_literal_keys.rs
// vybe-test-mode: compile

package main
type node struct { id int }
func main() { _ = []node{{id: 1}, {id: 2}} }

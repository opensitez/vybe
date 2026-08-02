// vybe-test: go/composite_literal_keys/struct_with_keyed_map_and_array_fields
// origin: languages/go/tests/go/test_composite_literal_keys.rs
// vybe-test-mode: compile

package main
type bundle struct { tags map[string]int
ids [2]int }
func main() { _ = bundle{tags: map[string]int{"x": 1}, ids: [2]int{1: 9}} }

// vybe-test: go/composite_literal_keys/array_of_maps_keyed_inner_entries
// origin: languages/go/tests/go/test_composite_literal_keys.rs
// vybe-test-mode: compile

package main
func main() { _ = [2]map[string]int{{"a": 1}, {"b": 2}} }

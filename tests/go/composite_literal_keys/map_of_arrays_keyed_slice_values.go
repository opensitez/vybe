// vybe-test: go/composite_literal_keys/map_of_arrays_keyed_slice_values
// origin: languages/go/tests/go/test_composite_literal_keys.rs
// vybe-test-mode: compile

package main
func main() { _ = map[string][3]int{"row": {0: 1, 2: 3}} }

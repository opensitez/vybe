// vybe-test: go/maps_patterns_extra/map_value_array_index_compile
// origin: languages/go/tests/go/test_maps_patterns_extra.rs
// vybe-test-mode: compile

package main
func main() { values := map[string][2]int{"a": [2]int{1, 2}}
_ = values["a"][1] }

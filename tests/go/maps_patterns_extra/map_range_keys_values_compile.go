// vybe-test: go/maps_patterns_extra/map_range_keys_values_compile
// origin: languages/go/tests/go/test_maps_patterns_extra.rs
// vybe-test-mode: compile

package main
func main() { values := map[string]int{"a": 1}
for key, value := range values { _, _ = key, value } }

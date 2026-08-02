// vybe-test: go/maps_patterns_extra/map_value_increment_compile
// origin: languages/go/tests/go/test_maps_patterns_extra.rs
// vybe-test-mode: compile

package main
func main() { values := map[string]int{"a": 1}
values["a"] = values["a"] + 1 }

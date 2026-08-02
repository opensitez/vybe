// vybe-test: go/maps_patterns_extra/map_lookup_missing_compile
// origin: languages/go/tests/go/test_maps_patterns_extra.rs
// vybe-test-mode: compile

package main
func main() { values := map[string]int{}
_ = values["missing"] }

// vybe-test: go/maps_patterns_extra/map_compare_to_nil_compile
// origin: languages/go/tests/go/test_maps_patterns_extra.rs
// vybe-test-mode: compile

package main
func main() { var values map[string]int
_ = (values == nil) }

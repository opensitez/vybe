// vybe-test: go/maps_patterns_extra/nil_map_delete_compile
// origin: languages/go/tests/go/test_maps_patterns_extra.rs
// vybe-test-mode: compile

package main
func main() { var values map[string]int
delete(values, "a") }

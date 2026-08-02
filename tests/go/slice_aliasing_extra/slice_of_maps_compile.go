// vybe-test: go/slice_aliasing_extra/slice_of_maps_compile
// origin: languages/go/tests/go/test_slice_aliasing_extra.rs
// vybe-test-mode: compile

package main
func main() { _ = []map[string]int{{"a": 1}} }

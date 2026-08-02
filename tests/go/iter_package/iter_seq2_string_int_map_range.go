// vybe-test: go/iter_package/iter_seq2_string_int_map_range
// origin: languages/go/tests/go/test_iter_package.rs
// vybe-test-mode: compile

package main
import "maps"
func main() { for k, v := range maps.All(map[string]int{"x": 1}) { _, _ = k, v } }

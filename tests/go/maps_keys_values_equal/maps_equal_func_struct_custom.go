// vybe-test: go/maps_keys_values_equal/maps_equal_func_struct_custom
// origin: languages/go/tests/go/test_maps_keys_values_equal.rs
// vybe-test-mode: compile

package main
import "maps"
type P struct { N int }
func main() { a := map[string]P{"a": {1}}
b := map[string]P{"a": {1}}
_ = maps.EqualFunc(a, b, func(x, y P) bool { return x.N == y.N }) }

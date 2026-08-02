// vybe-test: go/slices_maps_stdlib/maps_deletefunc_struct_values
// origin: languages/go/tests/go/test_slices_maps_stdlib.rs
// vybe-test-mode: compile

package main
import "maps"
type Pair struct { N int }
func main() { m := map[string]Pair{"a": {N: 1}}
maps.DeleteFunc(m, func(k string, v Pair) bool { return v.N == 0 }) }

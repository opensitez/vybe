// vybe-test: go/maps_patterns_extra/map_range_count_entries_runtime
// origin: languages/go/tests/go/test_maps_patterns_extra.rs

package main
import "fmt"
func main() { values := map[string]int{"a": 1, "b": 2, "c": 3}
count := 0
for range values { count++ }
fmt.Println(count)
}

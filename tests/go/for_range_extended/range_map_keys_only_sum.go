// vybe-test: go/for_range_extended/range_map_keys_only_sum
// origin: languages/go/tests/go/test_for_range_extended.rs

package main
import "fmt"
func main() { total := 0
for k := range map[string]int{"a": 1, "b": 2, "c": 3} { total += len(k) }
fmt.Println(total) }

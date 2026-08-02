// vybe-test: go/blank_identifier_extended/blank_range_map_discard_value
// origin: languages/go/tests/go/test_blank_identifier_extended.rs

package main
import "fmt"
func main() { keys := 0
for k := range map[string]int{"x": 1, "y": 2} { keys += len(k) }
fmt.Println(keys) }

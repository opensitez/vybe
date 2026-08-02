// vybe-test: go/for_range_extended/range_map_key_only_break_early
// origin: languages/go/tests/go/test_for_range_extended.rs

package main
import "fmt"
func main() { count := 0
for range map[string]int{"x": 1, "y": 2} { count++
break }
fmt.Println(count) }

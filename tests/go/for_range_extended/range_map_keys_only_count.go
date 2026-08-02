// vybe-test: go/for_range_extended/range_map_keys_only_count
// origin: languages/go/tests/go/test_for_range_extended.rs

package main
import "fmt"
func main() { count := 0
for range map[int]bool{1: true, 2: false, 3: true} { count++ }
fmt.Println(count) }

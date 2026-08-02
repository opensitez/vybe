// vybe-test: go/for_range_extended/range_nil_map_runtime_zero_iters
// origin: languages/go/tests/go/test_for_range_extended.rs

package main
import "fmt"
func main() { var m map[int]int
count := 0
for range m { count++ }
fmt.Println(count) }

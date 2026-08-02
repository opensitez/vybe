// vybe-test: go/range_iteration_extra/range_over_map_count_runtime
// origin: languages/go/tests/go/test_range_iteration_extra.rs

package main
import "fmt"
func main() { values := map[string]int{"a": 1, "b": 2}
count := 0
for range values { count++ }
fmt.Println(count)
}

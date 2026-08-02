// vybe-test: go/map_iteration_delete/nil_map_two_value_range_zero_iterations
// origin: languages/go/tests/go/test_map_iteration_delete.rs

package main
import "fmt"
func main() { var values map[string]int
count := 0
for key, value := range values { _, _ = key, value
count++ }
fmt.Println(count) }

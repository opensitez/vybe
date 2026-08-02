// vybe-test: go/range_iteration_extra/range_over_map_of_slices_runtime
// origin: languages/go/tests/go/test_range_iteration_extra.rs

package main
import "fmt"
func main() { values := map[string][]int{"a": []int{1, 2}, "b": []int{3}}
total := 0
for _, item := range values { total += len(item) }
fmt.Println(total)
}

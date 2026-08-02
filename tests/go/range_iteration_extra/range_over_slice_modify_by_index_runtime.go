// vybe-test: go/range_iteration_extra/range_over_slice_modify_by_index_runtime
// origin: languages/go/tests/go/test_range_iteration_extra.rs

package main
import "fmt"
func main() { values := []int{1, 2, 3}
for index := range values { values[index]++ }
fmt.Println(values[0])
fmt.Println(values[2])
}

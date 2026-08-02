// vybe-test: go/range_iteration_extra/range_over_struct_slice_field_runtime
// origin: languages/go/tests/go/test_range_iteration_extra.rs

package main
import "fmt"
type holder struct { values []int }
func main() { value := holder{values: []int{1, 2, 3}}
total := 0
for _, item := range value.values { total += item }
fmt.Println(total)
}

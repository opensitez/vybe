// vybe-test: go/range_iteration_extra/range_over_nil_slice_count_runtime
// origin: languages/go/tests/go/test_range_iteration_extra.rs

package main
import "fmt"
func main() { var values []int
count := 0
for range values { count++ }
fmt.Println(count)
}

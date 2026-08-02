// vybe-test: go/range_iteration_extra/range_over_empty_slice_runtime
// origin: languages/go/tests/go/test_range_iteration_extra.rs

package main
import "fmt"
func main() { count := 0
for range []int{} { count++ }
fmt.Println(count)
}

// vybe-test: go/range_iteration_extra/range_over_slice_of_strings_count_runtime
// origin: languages/go/tests/go/test_range_iteration_extra.rs

package main
import "fmt"
func main() { count := 0
for _, value := range []string{"go", "vybe"} { count += len(value) }
fmt.Println(count)
}

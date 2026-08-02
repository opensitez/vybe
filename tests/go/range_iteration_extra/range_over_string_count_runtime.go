// vybe-test: go/range_iteration_extra/range_over_string_count_runtime
// origin: languages/go/tests/go/test_range_iteration_extra.rs

package main
import "fmt"
func main() { count := 0
for range "vybe" { count++ }
fmt.Println(count)
}

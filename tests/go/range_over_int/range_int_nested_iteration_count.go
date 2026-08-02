// vybe-test: go/range_over_int/range_int_nested_iteration_count
// origin: languages/go/tests/go/test_range_over_int.rs

package main
import "fmt"
func main() { count := 0
for i := range 3 { for j := range 2 { count += i + j } }
fmt.Println(count) }

// vybe-test: go/range_iteration_extra/range_over_array_index_only_runtime
// origin: languages/go/tests/go/test_range_iteration_extra.rs

package main
import "fmt"
func main() { total := 0
for index := range [2]int{7, 8} { total += index }
fmt.Println(total)
}

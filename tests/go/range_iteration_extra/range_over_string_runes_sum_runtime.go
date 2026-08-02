// vybe-test: go/range_iteration_extra/range_over_string_runes_sum_runtime
// origin: languages/go/tests/go/test_range_iteration_extra.rs

package main
import "fmt"
func main() { total := 0
for _, value := range "AB" { total += int(value) }
fmt.Println(total)
}

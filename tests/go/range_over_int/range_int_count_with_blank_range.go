// vybe-test: go/range_over_int/range_int_count_with_blank_range
// origin: languages/go/tests/go/test_range_over_int.rs

package main
import "fmt"
func main() { count := 0
for range 6 { count++ }
fmt.Println(count) }

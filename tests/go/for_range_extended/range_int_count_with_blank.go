// vybe-test: go/for_range_extended/range_int_count_with_blank
// origin: languages/go/tests/go/test_for_range_extended.rs

package main
import "fmt"
func main() { count := 0
for range 8 { count++ }
fmt.Println(count) }

// vybe-test: go/for_range_extended/range_string_empty_zero_iters
// origin: languages/go/tests/go/test_for_range_extended.rs

package main
import "fmt"
func main() { count := 0
for range "" { count++ }
fmt.Println(count) }

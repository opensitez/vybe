// vybe-test: go/range_over_int/range_int_expression_bound
// origin: languages/go/tests/go/test_range_over_int.rs

package main
import "fmt"
func main() { total := 0
for i := range 2 + 3 { total += i }
fmt.Println(total) }

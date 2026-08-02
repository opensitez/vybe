// vybe-test: go/range_over_int/range_int_product_nonzero_indices
// origin: languages/go/tests/go/test_range_over_int.rs

package main
import "fmt"
func main() { total := 1
for i := range 4 { if i > 0 { total *= i } }
fmt.Println(total) }

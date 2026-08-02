// vybe-test: go/for_range_extended/range_int_nested_product_indices
// origin: languages/go/tests/go/test_for_range_extended.rs

package main
import "fmt"
func main() { product := 1
for i := range 3 { for j := range 2 { if i == 0 && j == 0 { continue }
product *= (i + 1) } }
fmt.Println(product) }

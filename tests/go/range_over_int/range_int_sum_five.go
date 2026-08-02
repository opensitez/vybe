// vybe-test: go/range_over_int/range_int_sum_five
// origin: languages/go/tests/go/test_range_over_int.rs

package main
import "fmt"
func main() { total := 0
for i := range 5 { total += i }
fmt.Println(total) }

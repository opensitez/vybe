// vybe-test: go/range_over_int/range_int_variable_bound
// origin: languages/go/tests/go/test_range_over_int.rs

package main
import "fmt"
func main() { n := 6
total := 0
for i := range n { total += i }
fmt.Println(total) }

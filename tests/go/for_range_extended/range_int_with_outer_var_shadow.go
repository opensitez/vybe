// vybe-test: go/for_range_extended/range_int_with_outer_var_shadow
// origin: languages/go/tests/go/test_for_range_extended.rs

package main
import "fmt"
func main() { bound := 4
total := 0
for i := range bound { total += i }
fmt.Println(total) }

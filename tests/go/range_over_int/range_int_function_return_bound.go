// vybe-test: go/range_over_int/range_int_function_return_bound
// origin: languages/go/tests/go/test_range_over_int.rs

package main
import "fmt"
func bound() int { return 4 }
func main() { total := 0
for i := range bound() { total += i }
fmt.Println(total) }

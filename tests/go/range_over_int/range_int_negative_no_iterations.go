// vybe-test: go/range_over_int/range_int_negative_no_iterations
// origin: languages/go/tests/go/test_range_over_int.rs

package main
import "fmt"
func main() { count := 0
for i := range -3 { count += i + 1 }
fmt.Println(count) }

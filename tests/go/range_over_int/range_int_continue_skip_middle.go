// vybe-test: go/range_over_int/range_int_continue_skip_middle
// origin: languages/go/tests/go/test_range_over_int.rs

package main
import "fmt"
func main() { total := 0
for i := range 4 { if i == 2 { continue }
total += i }
fmt.Println(total) }

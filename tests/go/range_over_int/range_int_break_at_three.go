// vybe-test: go/range_over_int/range_int_break_at_three
// origin: languages/go/tests/go/test_range_over_int.rs

package main
import "fmt"
func main() { total := 0
for i := range 8 { if i == 3 { break }
total += i }
fmt.Println(total) }

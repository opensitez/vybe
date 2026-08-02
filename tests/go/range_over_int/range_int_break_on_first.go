// vybe-test: go/range_over_int/range_int_break_on_first
// origin: languages/go/tests/go/test_range_over_int.rs

package main
import "fmt"
func main() { total := 0
for i := range 5 { total += i
break }
fmt.Println(total) }

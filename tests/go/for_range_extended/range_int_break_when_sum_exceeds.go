// vybe-test: go/for_range_extended/range_int_break_when_sum_exceeds
// origin: languages/go/tests/go/test_for_range_extended.rs

package main
import "fmt"
func main() { total := 0
for i := range 10 { total += i
if total > 5 { break } }
fmt.Println(total) }

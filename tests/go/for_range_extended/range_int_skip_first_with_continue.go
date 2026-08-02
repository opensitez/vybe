// vybe-test: go/for_range_extended/range_int_skip_first_with_continue
// origin: languages/go/tests/go/test_for_range_extended.rs

package main
import "fmt"
func main() { total := 0
for i := range 5 { if i == 0 { continue }
total += i }
fmt.Println(total) }

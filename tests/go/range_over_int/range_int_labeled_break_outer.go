// vybe-test: go/range_over_int/range_int_labeled_break_outer
// origin: languages/go/tests/go/test_range_over_int.rs

package main
import "fmt"
func main() { total := 0
outer: for i := range 4 { for j := range 3 { if i == 2 && j == 1 { break outer }
total += 1 } }
fmt.Println(total) }

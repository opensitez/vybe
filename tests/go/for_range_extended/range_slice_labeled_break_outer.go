// vybe-test: go/for_range_extended/range_slice_labeled_break_outer
// origin: languages/go/tests/go/test_for_range_extended.rs

package main
import "fmt"
func main() { total := 0
outer: for _, v := range []int{1, 2, 3, 4} { total += v
if v == 2 { break outer } }
fmt.Println(total) }

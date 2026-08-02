// vybe-test: go/for_range_extended/range_slice_continue_skip_negative
// origin: languages/go/tests/go/test_for_range_extended.rs

package main
import "fmt"
func main() { total := 0
for _, v := range []int{1, -1, 2, -2, 3} { if v < 0 { continue }
total += v }
fmt.Println(total) }

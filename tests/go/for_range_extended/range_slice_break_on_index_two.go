// vybe-test: go/for_range_extended/range_slice_break_on_index_two
// origin: languages/go/tests/go/test_for_range_extended.rs

package main
import "fmt"
func main() { count := 0
for i := range []int{9, 8, 7, 6} { count++
if i == 2 { break } }
fmt.Println(count) }

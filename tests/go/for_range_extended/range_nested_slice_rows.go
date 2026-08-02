// vybe-test: go/for_range_extended/range_nested_slice_rows
// origin: languages/go/tests/go/test_for_range_extended.rs

package main
import "fmt"
func main() { total := 0
for _, row := range [][]int{{1, 2}, {3}} { for _, v := range row { total += v } }
fmt.Println(total) }

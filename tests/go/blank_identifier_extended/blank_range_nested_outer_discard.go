// vybe-test: go/blank_identifier_extended/blank_range_nested_outer_discard
// origin: languages/go/tests/go/test_blank_identifier_extended.rs

package main
import "fmt"
func main() { total := 0
for _, row := range [][]int{{1, 2}, {3}} { for _, v := range row { total += v } }
fmt.Println(total) }

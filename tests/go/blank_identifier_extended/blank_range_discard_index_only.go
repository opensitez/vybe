// vybe-test: go/blank_identifier_extended/blank_range_discard_index_only
// origin: languages/go/tests/go/test_blank_identifier_extended.rs

package main
import "fmt"
func main() { total := 0
for _, v := range []int{2, 3, 4} { total += v }
fmt.Println(total) }

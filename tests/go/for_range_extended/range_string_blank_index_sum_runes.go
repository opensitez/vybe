// vybe-test: go/for_range_extended/range_string_blank_index_sum_runes
// origin: languages/go/tests/go/test_for_range_extended.rs

package main
import "fmt"
func main() { total := 0
for _, r := range "ab" { total += int(r) }
fmt.Println(total) }

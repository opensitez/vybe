// vybe-test: go/for_range_extended/range_string_rune_sum_ascii
// origin: languages/go/tests/go/test_for_range_extended.rs

package main
import "fmt"
func main() { total := 0
for _, r := range "Go" { total += int(r) }
fmt.Println(total) }

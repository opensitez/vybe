// vybe-test: go/blank_identifier_extended/blank_range_string_discard_index
// origin: languages/go/tests/go/test_blank_identifier_extended.rs

package main
import "fmt"
func main() { total := 0
for _, r := range "ab" { total += int(r) }
fmt.Println(total) }

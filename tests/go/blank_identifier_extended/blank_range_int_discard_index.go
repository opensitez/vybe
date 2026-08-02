// vybe-test: go/blank_identifier_extended/blank_range_int_discard_index
// origin: languages/go/tests/go/test_blank_identifier_extended.rs

package main
import "fmt"
func main() { count := 0
for range 4 { count++ }
fmt.Println(count) }

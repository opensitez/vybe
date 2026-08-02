// vybe-test: go/blank_identifier_extended/blank_range_discard_value_count_index
// origin: languages/go/tests/go/test_blank_identifier_extended.rs

package main
import "fmt"
func main() { count := 0
for range []string{"a", "b", "c"} { count++ }
fmt.Println(count) }

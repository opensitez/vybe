// vybe-test: go/for_range_extended/range_string_index_only_len
// origin: languages/go/tests/go/test_for_range_extended.rs

package main
import "fmt"
func main() { count := 0
for i := range "hello" { count = i + 1 }
fmt.Println(count) }

// vybe-test: go/for_range_extended/range_string_unicode_rune_count
// origin: languages/go/tests/go/test_for_range_extended.rs

package main
import "fmt"
func main() { count := 0
for range "日本" { count++ }
fmt.Println(count) }

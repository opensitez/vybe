// vybe-test: go/unicode_utf8/utf8_string_range_byte_indices
// origin: languages/go/tests/go/test_unicode_utf8.rs

package main
import "fmt"
func main() { first, second := -1, -1
step := 0
for i, _ := range "a世" { if step == 0 { first = i }
if step == 1 { second = i }
step++ }
fmt.Println(first)
fmt.Println(second) }
